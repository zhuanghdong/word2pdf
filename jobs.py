from __future__ import annotations

import asyncio
import logging
from dataclasses import dataclass, field
from datetime import UTC, datetime, timedelta
from uuid import uuid4

from config import Settings
from converter import ConversionService
from errors import (
    ConversionFailedError,
    JobNotFoundError,
    JobQueueFullError,
    ServiceError,
)
from schemas import JobDetails, JobStatus
from storage import WorkFiles, remove_work_dir


logger = logging.getLogger(__name__)


def _utcnow() -> datetime:
    return datetime.now(UTC)


@dataclass(slots=True)
class ConversionJob:
    job_id: str
    work: WorkFiles
    status: JobStatus = JobStatus.QUEUED
    created_at: datetime = field(default_factory=_utcnow)
    started_at: datetime | None = None
    finished_at: datetime | None = None
    expires_at: datetime | None = None
    error_code: str | None = None
    error_message: str | None = None

    def details(self) -> JobDetails:
        return JobDetails(
            job_id=self.job_id,
            status=self.status,
            created_at=self.created_at,
            started_at=self.started_at,
            finished_at=self.finished_at,
            expires_at=self.expires_at,
            error_code=self.error_code,
            error_message=self.error_message,
        )


class JobManager:
    def __init__(
        self,
        settings: Settings,
        converter: ConversionService,
    ) -> None:
        self._settings = settings
        self._converter = converter
        self._queue: asyncio.Queue[str | None] = asyncio.Queue(
            maxsize=settings.max_queued_jobs
        )
        self._jobs: dict[str, ConversionJob] = {}
        self._worker_task: asyncio.Task[None] | None = None
        self._cleanup_task: asyncio.Task[None] | None = None
        self._stop_event = asyncio.Event()
        self._stopping = False

    @property
    def is_running(self) -> bool:
        return (
            self._worker_task is not None
            and not self._worker_task.done()
            and not self._stopping
        )

    async def start(self) -> None:
        if self.is_running:
            return
        if self._worker_task is not None:
            self._queue = asyncio.Queue(maxsize=self._settings.max_queued_jobs)
        self._stopping = False
        self._stop_event.clear()
        self._worker_task = asyncio.create_task(
            self._worker_loop(), name="wordtopdf-job-worker"
        )
        self._cleanup_task = asyncio.create_task(
            self._cleanup_loop(), name="wordtopdf-job-cleanup"
        )

    async def stop(self) -> None:
        self._stopping = True
        self._stop_event.set()
        await self._queue.put(None)
        if self._worker_task is not None:
            await self._worker_task
        if self._cleanup_task is not None:
            await self._cleanup_task

        for job in list(self._jobs.values()):
            await remove_work_dir(job.work.work_dir)
        self._jobs.clear()

    async def submit(self, work: WorkFiles) -> ConversionJob:
        if not self.is_running:
            raise ConversionFailedError("The conversion worker is not running.")
        job = ConversionJob(job_id=uuid4().hex, work=work)
        self._jobs[job.job_id] = job
        try:
            self._queue.put_nowait(job.job_id)
        except asyncio.QueueFull as exc:
            self._jobs.pop(job.job_id, None)
            raise JobQueueFullError() from exc
        logger.info(
            "Conversion job queued",
            extra={"job_id": job.job_id, "size_bytes": work.size_bytes},
        )
        return job

    def get(self, job_id: str) -> ConversionJob:
        job = self._jobs.get(job_id)
        if job is None:
            raise JobNotFoundError()
        return job

    async def _worker_loop(self) -> None:
        while not self._stopping:
            job_id = await self._queue.get()
            try:
                if job_id is None or self._stopping:
                    return
                job = self._jobs.get(job_id)
                if job is None:
                    continue
                await self._run_job(job)
            finally:
                self._queue.task_done()

    async def _run_job(self, job: ConversionJob) -> None:
        job.status = JobStatus.RUNNING
        job.started_at = _utcnow()
        logger.info("Conversion job started", extra={"job_id": job.job_id})
        try:
            await self._converter.convert(job.work, correlation_id=job.job_id)
        except ServiceError as exc:
            job.status = JobStatus.FAILED
            job.error_code = exc.code
            job.error_message = exc.message
            internal_message = getattr(exc, "internal_message", None)
            logger.warning(
                "Conversion job failed",
                extra={
                    "job_id": job.job_id,
                    "error_code": exc.code,
                    "internal_message": internal_message,
                },
            )
        except Exception:
            job.status = JobStatus.FAILED
            job.error_code = "CONVERSION_FAILED"
            job.error_message = "Word 转 PDF 失败。"
            logger.exception(
                "Unexpected conversion job failure", extra={"job_id": job.job_id}
            )
        else:
            job.status = JobStatus.SUCCEEDED
            logger.info("Conversion job succeeded", extra={"job_id": job.job_id})
        finally:
            job.finished_at = _utcnow()
            job.expires_at = job.finished_at + timedelta(
                seconds=self._settings.job_ttl_seconds
            )

    async def _cleanup_loop(self) -> None:
        while not self._stopping:
            try:
                await asyncio.wait_for(
                    self._stop_event.wait(),
                    timeout=self._settings.cleanup_interval_seconds,
                )
            except TimeoutError:
                await self._expire_jobs()

    async def _expire_jobs(self) -> None:
        now = _utcnow()
        purge_before = now - timedelta(seconds=self._settings.job_ttl_seconds)
        for job_id, job in list(self._jobs.items()):
            if (
                job.status in {JobStatus.SUCCEEDED, JobStatus.FAILED}
                and job.expires_at is not None
                and job.expires_at <= now
            ):
                await remove_work_dir(job.work.work_dir)
                job.status = JobStatus.EXPIRED
            elif (
                job.status is JobStatus.EXPIRED
                and job.expires_at is not None
                and job.expires_at <= purge_before
            ):
                self._jobs.pop(job_id, None)

import asyncio
import logging
import multiprocessing
import shutil
from contextlib import asynccontextmanager
from pathlib import Path
from uuid import uuid4

import uvicorn
from fastapi import FastAPI, File, Query, Request, UploadFile
from fastapi.responses import FileResponse, JSONResponse
from starlette.background import BackgroundTask

from config import Settings, ensure_supported_platform, runtime_info
from converter import ConversionService
from errors import (
    ConversionFailedError,
    JobNotFoundError,
    JobNotReadyError,
    ServiceError,
)
from jobs import JobManager
from schemas import JobDetails, JobStatus, JobSubmitted
from storage import WorkFiles, remove_work_dir, save_upload


logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s %(levelname)s %(name)s %(message)s",
)
logger = logging.getLogger(__name__)

settings = Settings.from_env()
converter = ConversionService(settings)
job_manager = JobManager(settings, converter)


def _check_temp_directory(temp_root: Path) -> None:
    temp_root.mkdir(parents=True, exist_ok=True)
    probe_path = temp_root / f".health-{uuid4().hex}"
    probe_path.write_bytes(b"ok")
    probe_path.unlink()


@asynccontextmanager
async def lifespan(_: FastAPI):
    ensure_supported_platform()
    await asyncio.to_thread(_check_temp_directory, settings.temp_root)
    logger.info("Starting Word to PDF service", extra=runtime_info())
    await job_manager.start()
    try:
        yield
    finally:
        await job_manager.stop()


app = FastAPI(title="Word to PDF Service", lifespan=lifespan)


@app.middleware("http")
async def request_context(request: Request, call_next):
    request_id = request.headers.get("X-Request-ID") or uuid4().hex
    request.state.request_id = request_id
    try:
        response = await call_next(request)
    except Exception:
        logger.exception(
            "Unhandled request error",
            extra={"request_id": request_id, "path": request.url.path},
        )
        raise
    response.headers["X-Request-ID"] = request_id
    return response


@app.exception_handler(ServiceError)
async def service_error_handler(_: Request, exc: ServiceError) -> JSONResponse:
    internal_message = getattr(exc, "internal_message", None)
    if internal_message:
        logger.error(
            "Service operation failed",
            extra={"error_code": exc.code, "internal_message": internal_message},
        )
    return JSONResponse(
        status_code=exc.status_code,
        content={"detail": {"code": exc.code, "message": exc.message}},
    )


async def _receive_upload(file: UploadFile) -> WorkFiles:
    try:
        return await save_upload(file, settings)
    finally:
        await file.close()


@app.get("/health")
def health() -> dict[str, str]:
    return {"status": "ok"}


@app.get("/health/ready")
async def readiness(
    deep: bool = Query(
        default=False,
        description="是否启动 Word COM 做深度检查；检查期间会占用转换锁。",
    ),
) -> dict[str, object]:
    try:
        await asyncio.to_thread(_check_temp_directory, settings.temp_root)
        if not job_manager.is_running:
            raise ConversionFailedError("The background worker is not running.")
        if deep:
            await converter.probe()
    except ServiceError:
        raise
    except Exception as exc:
        raise ConversionFailedError(str(exc)) from exc
    return {
        "status": "ready",
        "worker": "running",
        "word_com": "available" if deep else "not_checked",
        "runtime": runtime_info(),
    }


@app.post("/convert")
async def convert(request: Request, file: UploadFile = File(...)) -> FileResponse:
    work = await _receive_upload(file)
    try:
        await converter.convert(work, correlation_id=request.state.request_id)
        return FileResponse(
            path=work.output_path,
            filename=work.download_filename,
            media_type="application/pdf",
            background=BackgroundTask(
                shutil.rmtree, work.work_dir, ignore_errors=True
            ),
        )
    except BaseException:
        await remove_work_dir(work.work_dir)
        raise


@app.post("/convert/jobs", response_model=JobSubmitted, status_code=202)
async def submit_conversion_job(file: UploadFile = File(...)) -> JobSubmitted:
    work = await _receive_upload(file)
    try:
        job = await job_manager.submit(work)
    except BaseException:
        await remove_work_dir(work.work_dir)
        raise
    return JobSubmitted(
        job_id=job.job_id,
        status=job.status,
        status_url=f"/convert/jobs/{job.job_id}",
        download_url=f"/convert/jobs/{job.job_id}/download",
    )


@app.get("/convert/jobs/{job_id}", response_model=JobDetails)
async def conversion_job_status(job_id: str) -> JobDetails:
    return job_manager.get(job_id).details()


@app.get("/convert/jobs/{job_id}/download")
async def download_conversion_result(job_id: str) -> FileResponse:
    job = job_manager.get(job_id)
    if job.status is JobStatus.EXPIRED:
        raise JobNotFoundError()
    if job.status is JobStatus.FAILED:
        raise ServiceError(
            job.error_code or "CONVERSION_FAILED",
            job.error_message or "Word 转 PDF 失败。",
            409,
        )
    if job.status is not JobStatus.SUCCEEDED:
        raise JobNotReadyError()
    if not job.work.output_path.is_file():
        raise ConversionFailedError("Job output file is missing.")
    return FileResponse(
        path=job.work.output_path,
        filename=job.work.download_filename,
        media_type="application/pdf",
    )


if __name__ == "__main__":
    multiprocessing.freeze_support()
    uvicorn.run("main:app", host="0.0.0.0", port=8000, reload=False, workers=1)
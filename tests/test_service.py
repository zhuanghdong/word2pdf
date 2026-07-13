from __future__ import annotations

import asyncio
import tempfile
import unittest
from io import BytesIO
from pathlib import Path

from fastapi import UploadFile

from config import Settings
from errors import ConversionTimeoutError, UnsupportedFileTypeError, UploadTooLargeError
from jobs import JobManager
from schemas import JobStatus
from storage import WorkFiles, save_upload, validate_extension


def make_settings(temp_root: Path, max_upload_bytes: int = 1024) -> Settings:
    return Settings(
        conversion_timeout_seconds=5,
        result_queue_timeout_seconds=1,
        max_upload_bytes=max_upload_bytes,
        upload_chunk_bytes=4,
        job_ttl_seconds=60,
        cleanup_interval_seconds=60,
        max_queued_jobs=100,
        temp_root=temp_root,
        aggressive_word_cleanup=False,
    )


class SuccessfulConverter:
    async def convert(
        self, work: WorkFiles, correlation_id: str | None = None
    ) -> None:
        work.output_path.write_bytes(b"%PDF-test")


class TimeoutConverter:
    async def convert(
        self, work: WorkFiles, correlation_id: str | None = None
    ) -> None:
        raise ConversionTimeoutError()


class StorageTests(unittest.IsolatedAsyncioTestCase):
    def test_extension_validation(self) -> None:
        self.assertEqual(validate_extension("report.DOCX"), ".docx")
        with self.assertRaises(UnsupportedFileTypeError):
            validate_extension("report.pdf")

    async def test_upload_is_written_in_chunks(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            upload = UploadFile(file=BytesIO(b"12345678"), filename="report.docx")
            work = await save_upload(upload, make_settings(root))
            self.assertEqual(work.input_path.read_bytes(), b"12345678")
            self.assertEqual(work.size_bytes, 8)
            self.assertEqual(work.download_filename, "report.pdf")

    async def test_oversized_upload_is_rejected_and_cleaned(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            upload = UploadFile(file=BytesIO(b"12345"), filename="large.docx")
            with self.assertRaises(UploadTooLargeError):
                await save_upload(upload, make_settings(root, max_upload_bytes=4))
            self.assertEqual(list(root.iterdir()), [])


class JobManagerTests(unittest.IsolatedAsyncioTestCase):
    async def _wait_until_finished(self, manager: JobManager, job_id: str) -> None:
        for _ in range(100):
            if manager.get(job_id).status in {
                JobStatus.SUCCEEDED,
                JobStatus.FAILED,
            }:
                return
            await asyncio.sleep(0.01)
        self.fail("Job did not finish in time.")

    async def test_successful_job_status_flow(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            work_dir = root / "job"
            work_dir.mkdir()
            work = WorkFiles(
                work_dir=work_dir,
                input_path=work_dir / "source.docx",
                output_path=work_dir / "result.pdf",
                original_filename="report.docx",
                size_bytes=4,
            )
            manager = JobManager(make_settings(root), SuccessfulConverter())  # type: ignore[arg-type]
            await manager.start()
            job = await manager.submit(work)
            await self._wait_until_finished(manager, job.job_id)
            self.assertEqual(manager.get(job.job_id).status, JobStatus.SUCCEEDED)
            self.assertTrue(work.output_path.exists())
            await manager.stop()
            self.assertFalse(work_dir.exists())

    async def test_service_error_is_recorded_on_job(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            work_dir = root / "job"
            work_dir.mkdir()
            work = WorkFiles(
                work_dir=work_dir,
                input_path=work_dir / "source.docx",
                output_path=work_dir / "result.pdf",
                original_filename="report.docx",
                size_bytes=4,
            )
            manager = JobManager(make_settings(root), TimeoutConverter())  # type: ignore[arg-type]
            await manager.start()
            job = await manager.submit(work)
            await self._wait_until_finished(manager, job.job_id)
            failed_job = manager.get(job.job_id)
            self.assertEqual(failed_job.status, JobStatus.FAILED)
            self.assertEqual(failed_job.error_code, "CONVERSION_TIMEOUT")
            await manager.stop()


if __name__ == "__main__":
    unittest.main()

from __future__ import annotations

from datetime import datetime
from enum import StrEnum

from pydantic import BaseModel


class JobStatus(StrEnum):
    QUEUED = "queued"
    RUNNING = "running"
    SUCCEEDED = "succeeded"
    FAILED = "failed"
    EXPIRED = "expired"


class JobSubmitted(BaseModel):
    job_id: str
    status: JobStatus
    status_url: str
    download_url: str


class JobDetails(BaseModel):
    job_id: str
    status: JobStatus
    created_at: datetime
    started_at: datetime | None = None
    finished_at: datetime | None = None
    expires_at: datetime | None = None
    error_code: str | None = None
    error_message: str | None = None

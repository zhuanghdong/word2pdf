from __future__ import annotations

import os
import platform
import struct
import tempfile
from dataclasses import dataclass
from pathlib import Path


def _env_int(name: str, default: int, minimum: int = 1) -> int:
    raw = os.getenv(name)
    if raw is None:
        return default
    try:
        value = int(raw)
    except ValueError as exc:
        raise ValueError(f"{name} must be an integer.") from exc
    if value < minimum:
        raise ValueError(f"{name} must be at least {minimum}.")
    return value


def _env_bool(name: str, default: bool) -> bool:
    raw = os.getenv(name)
    if raw is None:
        return default
    normalized = raw.strip().lower()
    if normalized in {"1", "true", "yes", "on"}:
        return True
    if normalized in {"0", "false", "no", "off"}:
        return False
    raise ValueError(f"{name} must be a boolean.")


@dataclass(frozen=True, slots=True)
class Settings:
    conversion_timeout_seconds: int
    result_queue_timeout_seconds: int
    max_upload_bytes: int
    upload_chunk_bytes: int
    job_ttl_seconds: int
    cleanup_interval_seconds: int
    max_queued_jobs: int
    temp_root: Path
    aggressive_word_cleanup: bool

    @classmethod
    def from_env(cls) -> "Settings":
        temp_root = Path(
            os.getenv(
                "WORDTOPDF_TEMP_ROOT",
                str(Path(tempfile.gettempdir()) / "wordtopdf"),
            )
        ).resolve()
        return cls(
            conversion_timeout_seconds=_env_int(
                "WORDTOPDF_CONVERSION_TIMEOUT_SECONDS", 120
            ),
            result_queue_timeout_seconds=_env_int(
                "WORDTOPDF_RESULT_QUEUE_TIMEOUT_SECONDS", 3
            ),
            max_upload_bytes=_env_int(
                "WORDTOPDF_MAX_UPLOAD_BYTES", 50 * 1024 * 1024
            ),
            upload_chunk_bytes=_env_int(
                "WORDTOPDF_UPLOAD_CHUNK_BYTES", 1024 * 1024
            ),
            job_ttl_seconds=_env_int("WORDTOPDF_JOB_TTL_SECONDS", 3600),
            cleanup_interval_seconds=_env_int(
                "WORDTOPDF_CLEANUP_INTERVAL_SECONDS", 60
            ),
            max_queued_jobs=_env_int("WORDTOPDF_MAX_QUEUED_JOBS", 100),
            temp_root=temp_root,
            aggressive_word_cleanup=_env_bool(
                "WORDTOPDF_AGGRESSIVE_WORD_CLEANUP", False
            ),
        )


def runtime_info() -> dict[str, str]:
    return {
        "platform": platform.system(),
        "platform_release": platform.release(),
        "python_version": platform.python_version(),
        "python_architecture": f"{struct.calcsize('P') * 8}-bit",
        "machine": platform.machine(),
    }


def ensure_supported_platform() -> None:
    if platform.system() != "Windows":
        raise RuntimeError("Word COM conversion is supported only on Windows.")

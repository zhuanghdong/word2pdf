from __future__ import annotations

import asyncio
import shutil
from dataclasses import dataclass
from pathlib import Path
from uuid import uuid4

from fastapi import UploadFile

from config import Settings
from errors import UnsupportedFileTypeError, UploadTooLargeError


WORD_EXTENSIONS = {".doc", ".docx", ".rtf"}


@dataclass(frozen=True, slots=True)
class WorkFiles:
    work_dir: Path
    input_path: Path
    output_path: Path
    original_filename: str
    size_bytes: int

    @property
    def download_filename(self) -> str:
        stem = Path(self.original_filename).stem or "document"
        safe_stem = "".join(
            character if character >= " " and character not in {'"', "\\", "/"} else "_"
            for character in stem
        )
        return f"{safe_stem}.pdf"


def validate_extension(filename: str) -> str:
    suffix = Path(filename).suffix.lower()
    if suffix not in WORD_EXTENSIONS:
        raise UnsupportedFileTypeError(WORD_EXTENSIONS)
    return suffix


def _open_for_write(path: Path):
    return path.open("wb")


async def save_upload(file: UploadFile, settings: Settings) -> WorkFiles:
    original_filename = file.filename or ""
    suffix = validate_extension(original_filename)
    work_dir = settings.temp_root / uuid4().hex
    input_path = work_dir / f"source{suffix}"
    output_path = work_dir / "result.pdf"
    await asyncio.to_thread(work_dir.mkdir, parents=True, exist_ok=False)

    size = 0
    try:
        output = await asyncio.to_thread(_open_for_write, input_path)
        try:
            while chunk := await file.read(settings.upload_chunk_bytes):
                size += len(chunk)
                if size > settings.max_upload_bytes:
                    raise UploadTooLargeError(settings.max_upload_bytes)
                await asyncio.to_thread(output.write, chunk)
            await asyncio.to_thread(output.flush)
        finally:
            await asyncio.to_thread(output.close)
    except BaseException:
        await asyncio.to_thread(shutil.rmtree, work_dir, True)
        raise

    return WorkFiles(
        work_dir=work_dir,
        input_path=input_path,
        output_path=output_path,
        original_filename=original_filename,
        size_bytes=size,
    )


async def remove_work_dir(work_dir: Path) -> None:
    await asyncio.to_thread(shutil.rmtree, work_dir, True)

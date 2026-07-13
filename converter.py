from __future__ import annotations

import asyncio
import logging
import multiprocessing
import queue
import shutil
import tempfile
import time
from pathlib import Path
from typing import Iterable

import psutil
import pythoncom
import win32com.client
import win32process

from config import Settings
from errors import (
    ConversionFailedError,
    ConversionTimeoutError,
    WordComUnavailableError,
)
from storage import WorkFiles


WORD_PROCESS_NAME = "WINWORD.EXE"
PDF_FORMAT = 17
logger = logging.getLogger(__name__)


def _best_effort_refresh_fields(document) -> None:
    """Refresh dynamic fields before PDF export.

    Broken cross-references/bookmarks inside the source document still need to
    be fixed in Word, but refreshing fields here avoids exporting stale values.
    """
    try:
        document.Fields.Update()
    except Exception:
        logger.exception("Failed to update Word fields before export")

    # TOC is not always covered by the general fields update path.
    try:
        for index in range(1, int(document.TablesOfContents.Count) + 1):
            document.TablesOfContents(index).Update()
    except Exception:
        logger.exception("Failed to update Word TOC before export")


def _winword_pids() -> set[int]:
    pids: set[int] = set()
    for proc in psutil.process_iter(["pid", "name"]):
        try:
            if (proc.info.get("name") or "").upper() == WORD_PROCESS_NAME:
                pids.add(int(proc.info["pid"]))
        except (psutil.NoSuchProcess, psutil.AccessDenied):
            continue
    return pids


def _is_word_process(proc: psutil.Process) -> bool:
    try:
        return (proc.name() or "").upper() == WORD_PROCESS_NAME
    except (psutil.NoSuchProcess, psutil.AccessDenied):
        return False


def _terminate_word_processes(pids: Iterable[int], reason: str) -> None:
    processes: list[psutil.Process] = []
    for pid in set(pids):
        try:
            process = psutil.Process(pid)
        except psutil.NoSuchProcess:
            continue
        if _is_word_process(process):
            processes.append(process)

    if not processes:
        return

    logger.warning(
        "Terminating managed Word processes",
        extra={"word_pids": [process.pid for process in processes], "reason": reason},
    )
    for process in processes:
        try:
            process.terminate()
        except (psutil.NoSuchProcess, psutil.AccessDenied):
            continue

    _, alive = psutil.wait_procs(processes, timeout=5)
    for process in alive:
        try:
            process.kill()
        except (psutil.NoSuchProcess, psutil.AccessDenied):
            continue


def _wait_for_word_exit_or_terminate(pid: int, reason: str) -> None:
    try:
        process = psutil.Process(pid)
    except psutil.NoSuchProcess:
        return
    if not _is_word_process(process):
        return
    try:
        process.wait(timeout=5)
    except psutil.TimeoutExpired:
        _terminate_word_processes({pid}, reason)
    except (psutil.NoSuchProcess, psutil.AccessDenied):
        return


def _write_pid_file(pid_file: Path, pid: int) -> None:
    pid_file.write_text(str(pid), encoding="ascii")


def _read_pid_file(pid_file: Path) -> set[int]:
    try:
        value = int(pid_file.read_text(encoding="ascii").strip())
    except (FileNotFoundError, OSError, ValueError):
        return set()
    return {value}


def _word_pid(word) -> int:
    bootstrap_document = None
    try:
        bootstrap_document = word.Documents.Add()
        _, pid = win32process.GetWindowThreadProcessId(
            int(word.ActiveWindow.Hwnd)
        )
        return int(pid)
    finally:
        if bootstrap_document is not None:
            try:
                bootstrap_document.Close(False)
            except Exception:
                logger.exception("Failed to close PID bootstrap document")


def _convert_with_word(input_path: Path, output_path: Path, pid_file: Path) -> None:
    pythoncom.CoInitialize()
    word = None
    document = None
    managed_pid: int | None = None
    try:
        try:
            word = win32com.client.DispatchEx("Word.Application")
        except Exception as exc:
            raise WordComUnavailableError() from exc
        word.Visible = False
        word.DisplayAlerts = 0
        managed_pid = _word_pid(word)
        _write_pid_file(pid_file, managed_pid)

        document = word.Documents.Open(
            str(input_path),
            ReadOnly=True,
            AddToRecentFiles=False,
            ConfirmConversions=False,
        )
        _best_effort_refresh_fields(document)
        document.SaveAs2(str(output_path), FileFormat=PDF_FORMAT)
    finally:
        if document is not None:
            try:
                document.Close(False)
            except Exception:
                logger.exception("Failed to close Word document")
        if word is not None:
            try:
                word.Quit()
            except Exception:
                logger.exception("Failed to quit Word")
        pythoncom.CoUninitialize()

        if managed_pid is not None:
            _wait_for_word_exit_or_terminate(managed_pid, "child-finalizer")


def _convert_worker(
    input_path: str,
    output_path: str,
    pid_file: str,
    result_queue,
) -> None:
    try:
        _convert_with_word(Path(input_path), Path(output_path), Path(pid_file))
        result_queue.put(("ok", None, None))
    except WordComUnavailableError as exc:
        result_queue.put(("service_error", exc.code, exc.message))
    except BaseException as exc:
        result_queue.put(
            ("error", "CONVERSION_FAILED", f"{type(exc).__name__}: {exc}")
        )


def _run_conversion_process(work: WorkFiles, settings: Settings) -> None:
    pid_file = work.work_dir / "winword.pid"
    before_pids = _winword_pids() if settings.aggressive_word_cleanup else set()
    context = multiprocessing.get_context("spawn")
    result_queue = context.Queue(maxsize=1)
    process = context.Process(
        target=_convert_worker,
        args=(
            str(work.input_path),
            str(work.output_path),
            str(pid_file),
            result_queue,
        ),
        name="wordtopdf-converter",
    )

    try:
        process.start()
        process.join(settings.conversion_timeout_seconds)

        if process.is_alive():
            process.terminate()
            process.join(5)
            if process.is_alive():
                process.kill()
                process.join(5)
            managed_pids = _read_pid_file(pid_file)
            if not managed_pids and settings.aggressive_word_cleanup:
                managed_pids = _winword_pids() - before_pids
            _terminate_word_processes(managed_pids, "conversion-timeout")
            raise ConversionTimeoutError()

        if process.exitcode != 0:
            _terminate_word_processes(
                _read_pid_file(pid_file), "converter-abnormal-exit"
            )
            raise ConversionFailedError(
                f"Converter process exited with code {process.exitcode}."
            )

        try:
            status, code, message = result_queue.get(
                timeout=settings.result_queue_timeout_seconds
            )
        except queue.Empty as exc:
            raise ConversionFailedError(
                "Converter process returned no result."
            ) from exc

        if status == "service_error" and code == "WORD_COM_UNAVAILABLE":
            raise WordComUnavailableError(message or "Microsoft Word COM 服务不可用。")
        if status != "ok":
            raise ConversionFailedError(message or "Unknown Word conversion error.")
        if not work.output_path.is_file() or work.output_path.stat().st_size == 0:
            raise ConversionFailedError("PDF output is missing or empty.")
    finally:
        if process.is_alive():
            process.kill()
            process.join(5)
        result_queue.close()
        result_queue.join_thread()


def _probe_word(pid_file: Path) -> None:
    pythoncom.CoInitialize()
    word = None
    managed_pid: int | None = None
    try:
        try:
            word = win32com.client.DispatchEx("Word.Application")
        except Exception as exc:
            raise WordComUnavailableError() from exc
        word.Visible = False
        word.DisplayAlerts = 0
        managed_pid = _word_pid(word)
        _write_pid_file(pid_file, managed_pid)
    finally:
        if word is not None:
            try:
                word.Quit()
            except Exception:
                logger.exception("Failed to quit Word after readiness probe")
        pythoncom.CoUninitialize()
        if managed_pid is not None:
            _wait_for_word_exit_or_terminate(managed_pid, "probe-finalizer")


def _probe_worker(pid_file: str, result_queue) -> None:
    try:
        _probe_word(Path(pid_file))
        result_queue.put(("ok", None))
    except BaseException as exc:
        result_queue.put(("error", f"{type(exc).__name__}: {exc}"))


def _run_probe_process(settings: Settings) -> None:
    work_dir = Path(
        tempfile.mkdtemp(prefix="probe-", dir=str(settings.temp_root))
    )
    pid_file = work_dir / "winword.pid"
    context = multiprocessing.get_context("spawn")
    result_queue = context.Queue(maxsize=1)
    process = context.Process(
        target=_probe_worker,
        args=(str(pid_file), result_queue),
        name="wordtopdf-readiness-probe",
    )
    try:
        process.start()
        process.join(min(settings.conversion_timeout_seconds, 30))
        if process.is_alive():
            process.terminate()
            process.join(5)
            if process.is_alive():
                process.kill()
                process.join(5)
            _terminate_word_processes(_read_pid_file(pid_file), "probe-timeout")
            raise WordComUnavailableError("Microsoft Word COM 就绪检查超时。")
        if process.exitcode != 0:
            _terminate_word_processes(
                _read_pid_file(pid_file), "probe-abnormal-exit"
            )
            raise WordComUnavailableError()
        try:
            status, message = result_queue.get(
                timeout=settings.result_queue_timeout_seconds
            )
        except queue.Empty as exc:
            raise WordComUnavailableError(
                "Microsoft Word COM 就绪检查未返回结果。"
            ) from exc
        if status != "ok":
            logger.error(
                "Word COM readiness probe failed",
                extra={"internal_message": message},
            )
            raise WordComUnavailableError()
    finally:
        if process.is_alive():
            process.kill()
            process.join(5)
        result_queue.close()
        result_queue.join_thread()
        shutil.rmtree(work_dir, ignore_errors=True)


class ConversionService:
    def __init__(self, settings: Settings) -> None:
        self._settings = settings
        self._lock = asyncio.Lock()

    async def convert(
        self,
        work: WorkFiles,
        correlation_id: str | None = None,
    ) -> None:
        started = time.monotonic()
        async with self._lock:
            logger.info(
                "Word conversion started",
                extra={
                    "work_dir": str(work.work_dir),
                    "size_bytes": work.size_bytes,
                    "suffix": work.input_path.suffix,
                    "correlation_id": correlation_id,
                },
            )
            try:
                await asyncio.to_thread(_run_conversion_process, work, self._settings)
            finally:
                logger.info(
                    "Word conversion finished",
                    extra={
                        "work_dir": str(work.work_dir),
                        "elapsed_seconds": round(time.monotonic() - started, 3),
                        "correlation_id": correlation_id,
                    },
                )

    async def probe(self) -> None:
        async with self._lock:
            await asyncio.to_thread(_run_probe_process, self._settings)

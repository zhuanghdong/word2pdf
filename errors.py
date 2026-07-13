from __future__ import annotations

from dataclasses import dataclass

from fastapi import HTTPException


@dataclass(slots=True)
class ServiceError(Exception):
    code: str
    message: str
    status_code: int = 500

    def __str__(self) -> str:
        return self.message


class UnsupportedFileTypeError(ServiceError):
    def __init__(self, supported: set[str]) -> None:
        super().__init__(
            "UNSUPPORTED_FILE_TYPE",
            f"仅支持以下文件类型：{', '.join(sorted(supported))}",
            400,
        )


class UploadTooLargeError(ServiceError):
    def __init__(self, max_bytes: int) -> None:
        super().__init__(
            "UPLOAD_TOO_LARGE",
            f"上传文件超过大小限制（最大 {max_bytes} 字节）。",
            413,
        )


class ConversionTimeoutError(ServiceError):
    def __init__(self) -> None:
        super().__init__("CONVERSION_TIMEOUT", "Word 转换超时。", 504)


class WordComUnavailableError(ServiceError):
    def __init__(self, message: str = "Microsoft Word COM 服务不可用。") -> None:
        super().__init__("WORD_COM_UNAVAILABLE", message, 503)


class ConversionFailedError(ServiceError):
    def __init__(self, internal_message: str | None = None) -> None:
        super().__init__("CONVERSION_FAILED", "Word 转 PDF 失败。", 500)
        self.internal_message = internal_message


class JobNotFoundError(ServiceError):
    def __init__(self) -> None:
        super().__init__("JOB_NOT_FOUND", "转换任务不存在或已过期。", 404)


class JobNotReadyError(ServiceError):
    def __init__(self) -> None:
        super().__init__("JOB_NOT_READY", "转换任务尚未完成。", 409)


class JobQueueFullError(ServiceError):
    def __init__(self) -> None:
        super().__init__("JOB_QUEUE_FULL", "转换队列已满，请稍后重试。", 429)


def as_http_exception(exc: ServiceError) -> HTTPException:
    return HTTPException(
        status_code=exc.status_code,
        detail={"code": exc.code, "message": exc.message},
    )

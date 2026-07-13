# Word 转 PDF 服务

基于 Windows、Microsoft Word COM 和 FastAPI 的 Word 转 PDF 服务，同时提供同步返回和异步任务接口。

## 环境要求

- Windows 10/11 或受支持的 Windows Server
- 已安装并激活 Microsoft Word
- Python 3.11+
- 建议使用独立 Windows 服务账户

安装并启动：

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
uvicorn main:app --host 0.0.0.0 --port 8000 --workers 1
```

也可以直接运行：

```powershell
python .\main.py
```

必须保持一个 Uvicorn worker。Word COM 转换在应用内部串行执行，多 worker 会产生互相独立的队列和任务状态。

## API

### 同步转换

`POST /convert` 接收字段名为 `file` 的 multipart 文件，转换完成后直接返回 PDF：

```powershell
curl.exe -X POST http://127.0.0.1:8000/convert `
  -F "file=@D:\docs\sample.docx" `
  --output sample.pdf
```

### 异步转换

提交任务：

```powershell
curl.exe -X POST http://127.0.0.1:8000/convert/jobs `
  -F "file=@D:\docs\sample.docx"
```

服务返回 `202 Accepted`：

```json
{
  "job_id": "任务ID",
  "status": "queued",
  "status_url": "/convert/jobs/任务ID",
  "download_url": "/convert/jobs/任务ID/download"
}
```

查询状态：

```powershell
curl.exe http://127.0.0.1:8000/convert/jobs/任务ID
```

任务状态包括 `queued`、`running`、`succeeded`、`failed` 和 `expired`。成功后下载：

```powershell
curl.exe http://127.0.0.1:8000/convert/jobs/任务ID/download `
  --output sample.pdf
```

异步任务状态保存在当前进程内存中，服务重启后不会恢复。结果默认保留一小时；需要持久化或多实例部署时，应将任务状态迁移到 Redis/数据库，并将文件迁移到共享或对象存储。

### 健康检查

- `GET /health`：轻量存活检查。
- `GET /health/ready`：检查临时目录和后台 worker。
- `GET /health/ready?deep=true`：额外启动一次 Word COM 探测；该检查会占用转换锁，不应高频调用。

## 文件与错误约定

- 支持 `.doc`、`.docx`、`.rtf`。
- 默认最大上传大小为 50 MiB。
- 错误响应使用稳定错误码：

```json
{
  "detail": {
    "code": "CONVERSION_TIMEOUT",
    "message": "Word 转换超时。"
  }
}
```

常见错误码包括 `UNSUPPORTED_FILE_TYPE`、`UPLOAD_TOO_LARGE`、`CONVERSION_TIMEOUT`、`WORD_COM_UNAVAILABLE`、`CONVERSION_FAILED`、`JOB_QUEUE_FULL`、`JOB_NOT_FOUND` 和 `JOB_NOT_READY`。

## 配置

通过环境变量配置：

| 环境变量 | 默认值 | 说明 |
| --- | ---: | --- |
| `WORDTOPDF_CONVERSION_TIMEOUT_SECONDS` | `120` | 单次转换超时 |
| `WORDTOPDF_RESULT_QUEUE_TIMEOUT_SECONDS` | `3` | 子进程结果队列读取超时 |
| `WORDTOPDF_MAX_UPLOAD_BYTES` | `52428800` | 最大上传字节数 |
| `WORDTOPDF_UPLOAD_CHUNK_BYTES` | `1048576` | 上传分块大小 |
| `WORDTOPDF_JOB_TTL_SECONDS` | `3600` | 异步结果保留时间 |
| `WORDTOPDF_CLEANUP_INTERVAL_SECONDS` | `60` | 过期任务清理间隔 |
| `WORDTOPDF_MAX_QUEUED_JOBS` | `100` | 等待中的异步任务上限 |
| `WORDTOPDF_TEMP_ROOT` | 系统临时目录下的 `wordtopdf` | 工作目录 |
| `WORDTOPDF_AGGRESSIVE_WORD_CLEANUP` | `false` | 无法取得精确 PID 时是否使用进程差集兜底 |

生产环境不建议开启激进清理；它可能终止同一账户下由其他程序同时启动的 Word 进程。

## Win32 / Win64

- 支持 32 位或 64 位 Python，只要安装与 Python 架构匹配的 pywin32 包。
- Word 是进程外 COM 服务，Office 与 Python 位数不一定必须相同，但同架构组合更容易部署和排障。
- 生产环境推荐 64 位 Windows、64 位 Python 和 64 位 Office。
- `/health/ready` 会返回 Python 架构和系统架构，便于诊断。

## 可靠性设计

- 每次转换运行在独立的 `spawn` 子进程中，COM 卡死时父进程可以超时终止。
- 使用 Word 窗口句柄解析本次 `WINWORD.EXE` 的精确 PID，只清理由本服务管理的进程。
- 默认不使用 Word 进程前后差集，避免误杀其他用户或程序的 Word。
- 使用带超时的队列 `get()` 读取子进程结果，不依赖 Windows 上不可靠的 `Queue.empty()`。
- 同步和异步接口共用同一个转换服务与异步锁，避免 Word COM 并发操作。
- 上传分块读取并在线程池写盘，超限后立即清理临时文件。
- 对外隐藏 COM 异常细节，详细原因只写服务日志。

## Windows 服务部署注意事项

- Microsoft 不建议在无人值守服务端自动化 Office。若业务必须使用该方案，应接受 Office 弹窗、首次启动配置或异常文档导致卡死的固有限制。
- 使用独立账户，不要与人工桌面会话共用账户。
- 首次部署时以该账户登录并手动打开 Word，完成激活、隐私提示和模板初始化。
- 配置 NSSM、Windows Service 或任务计划程序的失败恢复。
- 密码保护、损坏或需要交互确认的文档可能失败。

## 测试

运行不依赖 Office 的单元测试：

```powershell
python -m unittest discover -s tests -v
```

上线前还应使用实际安装的 Word 分别验证同步转换、异步转换、超时清理以及目标 Win32/Win64 组合。

import os
import threading
import time
from datetime import datetime, timedelta
from pathlib import Path
from typing import Optional
from uuid import uuid4

try:
    import boto3  # type: ignore
except Exception:  # pragma: no cover - boto3 may not be installed
    boto3 = None  # type: ignore

# Directory configuration
BASE_DIR = Path(__file__).resolve().parent
DEFAULT_STORAGE = BASE_DIR / "storage"
STORAGE_DIR = Path(os.getenv("STORAGE_DIR", DEFAULT_STORAGE))
STORAGE_DIR.mkdir(parents=True, exist_ok=True)

FILE_TTL_SECONDS = int(os.getenv("FILE_TTL_SECONDS", 60 * 60 * 24))
USE_S3 = os.getenv("USE_S3", "0") == "1"
S3_BUCKET = os.getenv("S3_BUCKET") if USE_S3 else None

if USE_S3 and boto3 is None:
    raise RuntimeError("USE_S3 is enabled but boto3 is not installed")

if USE_S3:
    s3_client = boto3.client("s3")  # type: ignore


def _save_local(content: bytes, ext: str) -> str:
    file_id = uuid4().hex
    filename = f"{file_id}.{ext}"
    file_path = STORAGE_DIR / filename
    with open(file_path, "wb") as f:
        f.write(content)
    return file_id


def _save_s3(content: bytes, ext: str) -> str:
    file_id = uuid4().hex
    filename = f"{file_id}.{ext}"
    s3_client.put_object(Bucket=S3_BUCKET, Key=filename, Body=content)  # type: ignore
    return file_id


def save_pdf(content: bytes) -> str:
    """Save PDF content and return its unique identifier."""
    if USE_S3:
        return _save_s3(content, "pdf")
    return _save_local(content, "pdf")


def save_excel(content: bytes) -> str:
    """Save Excel content and return its unique identifier."""
    if USE_S3:
        return _save_s3(content, "xlsx")
    return _save_local(content, "xlsx")


def get_file_path(file_id: str, ext: str) -> Path:
    return STORAGE_DIR / f"{file_id}.{ext}"


def generate_presigned_url(file_id: str, ext: str, expires: int = 3600) -> Optional[str]:
    if not USE_S3:
        return None
    key = f"{file_id}.{ext}"
    return s3_client.generate_presigned_url(
        "get_object", Params={"Bucket": S3_BUCKET, "Key": key}, ExpiresIn=expires
    )  # type: ignore


def cleanup_old_files() -> None:
    if USE_S3:
        return  # assume external lifecycle management
    cutoff = datetime.utcnow() - timedelta(seconds=FILE_TTL_SECONDS)
    for file in STORAGE_DIR.iterdir():
        if file.is_file():
            mtime = datetime.utcfromtimestamp(file.stat().st_mtime)
            if mtime < cutoff:
                try:
                    file.unlink()
                except OSError:
                    pass


def start_cleanup_thread(interval_seconds: int = 3600) -> None:
    def _runner() -> None:
        while True:
            cleanup_old_files()
            time.sleep(interval_seconds)

    thread = threading.Thread(target=_runner, daemon=True)
    thread.start()

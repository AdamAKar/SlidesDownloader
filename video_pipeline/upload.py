"""Utilities for handling video uploads from file pickers or drag-and-drop events."""

from pathlib import Path
from typing import Iterable, Optional
import logging

logger = logging.getLogger(__name__)

# Default upload directory used by the service.
UPLOAD_DIR = Path("tmp/uploads")


def ensure_directory(path: Path) -> Path:
    """Ensure the directory exists and return it."""
    path.mkdir(parents=True, exist_ok=True)
    return path


def save_upload_bytes(
    file_name: str,
    data: bytes,
    content_type: Optional[str] = None,
    allowed_types: Optional[Iterable[str]] = None,
    max_size_bytes: int = 500 * 1024 * 1024,
    upload_dir: Path = UPLOAD_DIR,
) -> Path:
    """
    Validate and persist uploaded bytes to disk.

    Args:
        file_name: Original file name from the picker or drag-and-drop payload.
        data: Raw file bytes.
        content_type: MIME type supplied by the client.
        allowed_types: Iterable of accepted MIME types. If None, all types are accepted.
        max_size_bytes: Maximum file size allowed. Defaults to 500MB.
        upload_dir: Destination folder for persisted uploads.

    Returns:
        Path to the saved file on disk.

    Raises:
        ValueError: If validation fails.
    """
    logger.debug("Received upload %s (%s bytes, type=%s)", file_name, len(data), content_type)

    if len(data) == 0:
        raise ValueError("Uploaded file is empty.")

    if len(data) > max_size_bytes:
        raise ValueError(f"Uploaded file exceeds maximum size of {max_size_bytes} bytes.")

    if allowed_types is not None and content_type not in set(allowed_types):
        raise ValueError(f"Content type {content_type} is not allowed.")

    ensure_directory(upload_dir)
    safe_name = Path(file_name).name
    destination = upload_dir / safe_name

    logger.info("Persisting upload to %s", destination)
    destination.write_bytes(data)
    return destination

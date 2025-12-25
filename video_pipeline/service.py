"""Core service orchestration for the video pipeline."""

from __future__ import annotations

import argparse
import json
import logging
from pathlib import Path
from typing import Dict, Iterable, Optional

from .metadata import extract_metadata
from .normalize import normalize_video, DEFAULT_OUTPUT_DIR
from .upload import save_upload_bytes, UPLOAD_DIR

logger = logging.getLogger(__name__)


class VideoPipelineService:
    """Process uploads through metadata extraction and normalization."""

    def __init__(
        self,
        allowed_types: Optional[Iterable[str]] = None,
        max_size_bytes: int = 500 * 1024 * 1024,
        upload_dir: Path = UPLOAD_DIR,
        normalized_dir: Path = DEFAULT_OUTPUT_DIR,
    ) -> None:
        self.allowed_types = set(allowed_types) if allowed_types else None
        self.max_size_bytes = max_size_bytes
        self.upload_dir = upload_dir
        self.normalized_dir = normalized_dir

    def process_upload(self, file_name: str, data: bytes, content_type: Optional[str]) -> Dict:
        """Persist upload, extract metadata, run normalization, and return status."""
        upload_path = save_upload_bytes(
            file_name=file_name,
            data=data,
            content_type=content_type,
            allowed_types=self.allowed_types,
            max_size_bytes=self.max_size_bytes,
            upload_dir=self.upload_dir,
        )

        metadata, sidecar_path = extract_metadata(upload_path)
        normalized_path = normalize_video(upload_path, output_dir=self.normalized_dir)

        return {
            "upload_path": str(upload_path),
            "metadata_path": str(sidecar_path),
            "normalized_path": str(normalized_path),
            "metadata": metadata,
        }


def create_app(service: Optional[VideoPipelineService] = None):
    """Create a FastAPI app exposing an upload endpoint."""
    from fastapi import FastAPI, File, UploadFile

    pipeline = service or VideoPipelineService(allowed_types={"video/mp4", "video/quicktime", "video/x-matroska"})
    app = FastAPI()

    @app.post("/upload")
    async def upload_endpoint(file: UploadFile = File(...)):
        data = await file.read()
        result = pipeline.process_upload(file.filename, data, file.content_type)
        return result

    @app.get("/health")
    async def healthcheck():
        return {"status": "ok"}

    return app


def main():
    """CLI entrypoint supporting local uploads without HTTP."""
    parser = argparse.ArgumentParser(description="Run video pipeline on a file upload.")
    parser.add_argument("file", type=Path, help="Path to a video file to process")
    parser.add_argument("--content-type", dest="content_type", default="video/mp4", help="MIME type of the file")
    parser.add_argument(
        "--allowed-types",
        nargs="*",
        default=["video/mp4", "video/quicktime", "video/x-matroska"],
        help="List of allowed MIME types",
    )
    parser.add_argument("--max-size", dest="max_size", type=int, default=500 * 1024 * 1024, help="Maximum size in bytes")
    parser.add_argument("--serve", action="store_true", help="Run the FastAPI app instead of processing locally")
    args = parser.parse_args()

    if args.serve:
        from uvicorn import run

        app = create_app(
            VideoPipelineService(
                allowed_types=args.allowed_types,
                max_size_bytes=args.max_size,
            )
        )
        run(app, host="0.0.0.0", port=8000)
        return

    file_path = args.file
    data = file_path.read_bytes()
    service = VideoPipelineService(
        allowed_types=args.allowed_types,
        max_size_bytes=args.max_size,
    )
    result = service.process_upload(file_path.name, data, args.content_type)
    print(json.dumps(result, indent=2))


if __name__ == "__main__":
    main()

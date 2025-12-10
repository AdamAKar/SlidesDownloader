"""Normalization pipeline for transcoding uploads to a consistent format."""

import logging
import subprocess
from pathlib import Path
from typing import Optional

logger = logging.getLogger(__name__)

DEFAULT_OUTPUT_DIR = Path("tmp/normalized")


def normalize_video(
    input_path: Path,
    output_dir: Path = DEFAULT_OUTPUT_DIR,
    container: str = "mp4",
    video_codec: str = "libx264",
    audio_codec: str = "aac",
    overwrite: bool = True,
) -> Path:
    """
    Transcode a video into the configured working format using ffmpeg.

    Args:
        input_path: Source video file.
        output_dir: Destination directory for normalized media.
        container: Target container/extension.
        video_codec: ffmpeg video codec.
        audio_codec: ffmpeg audio codec.
        overwrite: Whether to overwrite existing outputs.

    Returns:
        Path to the normalized video.

    Raises:
        RuntimeError: If ffmpeg exits with a non-zero status.
    """
    output_dir.mkdir(parents=True, exist_ok=True)
    output_name = f"{input_path.stem}.{container}"
    output_path = output_dir / output_name

    command = [
        "ffmpeg",
        "-y" if overwrite else "-n",
        "-i",
        str(input_path),
        "-c:v",
        video_codec,
        "-c:a",
        audio_codec,
        str(output_path),
    ]

    logger.info("Starting normalization for %s", input_path)
    try:
        subprocess.run(command, check=True, capture_output=True, text=True)
    except subprocess.CalledProcessError as exc:  # pragma: no cover - defensive logging branch
        logger.error("ffmpeg failed (%s): %s", exc.returncode, exc.stderr)
        raise RuntimeError("Normalization failed") from exc

    logger.debug("Normalization completed: %s", output_path)
    return output_path

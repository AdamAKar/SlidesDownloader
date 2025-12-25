"""Metadata extraction helpers powered by ffprobe."""

from __future__ import annotations

import json
import logging
import subprocess
from pathlib import Path
from typing import Dict, Optional, Tuple

logger = logging.getLogger(__name__)


def _run_ffprobe(video_path: Path) -> Dict:
    """Invoke ffprobe and return the parsed JSON output."""
    command = [
        "ffprobe",
        "-v",
        "quiet",
        "-print_format",
        "json",
        "-show_streams",
        "-show_format",
        str(video_path),
    ]
    logger.debug("Running ffprobe: %s", " ".join(command))
    result = subprocess.run(command, check=True, capture_output=True, text=True)
    return json.loads(result.stdout)


def parse_metadata(probe_output: Dict) -> Dict:
    """Extract key metadata fields from a raw ffprobe response."""
    video_stream = next((s for s in probe_output.get("streams", []) if s.get("codec_type") == "video"), {})
    format_info = probe_output.get("format", {})

    width = video_stream.get("width")
    height = video_stream.get("height")
    codec = video_stream.get("codec_name")
    duration = format_info.get("duration") or video_stream.get("duration")

    try:
        parsed_duration: Optional[float] = float(duration) if duration is not None else None
    except (TypeError, ValueError):
        parsed_duration = None

    metadata = {
        "width": width,
        "height": height,
        "codec": codec,
        "duration": parsed_duration,
    }
    logger.debug("Parsed metadata: %s", metadata)
    return metadata


def extract_metadata(video_path: Path) -> Tuple[Dict, Path]:
    """
    Extract metadata and persist a JSON sidecar next to the video.

    Returns:
        A tuple of (metadata dictionary, path to sidecar JSON file).
    """
    probe_output = _run_ffprobe(video_path)
    metadata = parse_metadata(probe_output)

    sidecar_path = video_path.with_suffix(video_path.suffix + ".metadata.json")
    logger.info("Writing metadata sidecar to %s", sidecar_path)
    sidecar_path.write_text(json.dumps(metadata, indent=2))
    return metadata, sidecar_path

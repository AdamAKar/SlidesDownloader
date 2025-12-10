import json
from pathlib import Path
from typing import Any

import pytest

from video_pipeline import metadata


def test_parse_metadata_extracts_fields():
    probe_output = {
        "streams": [
            {"codec_type": "audio", "codec_name": "aac"},
            {"codec_type": "video", "codec_name": "h264", "width": 1920, "height": 1080, "duration": "10.5"},
        ],
        "format": {"duration": "10.50"},
    }

    parsed = metadata.parse_metadata(probe_output)
    assert parsed["width"] == 1920
    assert parsed["height"] == 1080
    assert parsed["codec"] == "h264"
    assert parsed["duration"] == 10.5


def test_extract_metadata_writes_sidecar(monkeypatch, tmp_path: Path):
    video_path = tmp_path / "clip.mov"
    video_path.write_bytes(b"fake data")

    def fake_run_ffprobe(path: Path) -> Any:
        return {
            "streams": [
                {"codec_type": "video", "codec_name": "vp9", "width": 640, "height": 360},
            ],
            "format": {"duration": "2"},
        }

    monkeypatch.setattr(metadata, "_run_ffprobe", fake_run_ffprobe)

    parsed, sidecar = metadata.extract_metadata(video_path)

    assert parsed == {"width": 640, "height": 360, "codec": "vp9", "duration": 2.0}
    assert sidecar.exists()
    saved = json.loads(sidecar.read_text())
    assert saved == parsed

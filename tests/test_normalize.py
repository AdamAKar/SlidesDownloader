from pathlib import Path
from typing import List
import subprocess

import pytest

from video_pipeline import normalize


def test_normalize_video_invokes_ffmpeg(monkeypatch, tmp_path: Path):
    input_path = tmp_path / "video.mkv"
    input_path.write_bytes(b"fake")
    commands: List[List[str]] = []

    def fake_run(command, check, capture_output, text):  # noqa: ANN001
        commands.append(command)
        return None

    monkeypatch.setattr(normalize.subprocess, "run", fake_run)

    output = normalize.normalize_video(input_path, output_dir=tmp_path)

    assert output == tmp_path / "video.mp4"
    assert commands
    invoked = commands[0]
    assert invoked[:2] == ["ffmpeg", "-y"]
    assert invoked[2:4] == ["-i", str(input_path)]
    assert invoked[-3:] == ["-c:a", "aac", str(output)]


def test_normalize_video_errors(monkeypatch, tmp_path: Path):
    input_path = tmp_path / "video.mov"
    input_path.write_bytes(b"fake")

    class FakeCalledProcessError(subprocess.CalledProcessError):
        def __init__(self):
            super().__init__(1, cmd="ffmpeg", stderr="boom")

    def fake_run(command, check, capture_output, text):  # noqa: ANN001
        raise FakeCalledProcessError()

    monkeypatch.setattr(normalize.subprocess, "run", fake_run)

    with pytest.raises(RuntimeError):
        normalize.normalize_video(input_path, output_dir=tmp_path)

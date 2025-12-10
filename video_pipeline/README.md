# Video Pipeline

This package implements a lightweight upload → metadata → normalization workflow for video assets.

## Components
- `upload.py`: Validates uploads from file pickers or drag-and-drop events and writes them to `tmp/uploads/`.
- `metadata.py`: Uses `ffprobe` to extract resolution, codec, and duration, persisting a JSON sidecar next to the upload.
- `normalize.py`: Invokes `ffmpeg` to transcode to MP4/H.264 + AAC in `tmp/normalized/`.
- `service.py`: Orchestrates the flow and exposes CLI + FastAPI entrypoints for manual testing and integrations.

## Setup
Install dependencies (FastAPI for HTTP entrypoints and pytest for tests):

```bash
pip install fastapi uvicorn pytest
```

`ffmpeg`/`ffprobe` must be available on your PATH for runtime processing.

## Usage

### Run the FastAPI server
```bash
python -m video_pipeline.service --serve
```
Then POST `multipart/form-data` with a `file` field to `http://localhost:8000/upload` using either a file picker or drag-and-drop-capable UI.

### CLI processing
Process a local file without starting the server:
```bash
python -m video_pipeline.service path/to/video.mp4 --content-type video/mp4
```

### Configuration
- Allowed MIME types: configure via `--allowed-types` or the `VideoPipelineService` constructor.
- Max size: configure via `--max-size` (bytes). Default is 500MB.
- Output locations: uploads under `tmp/uploads/`; normalized outputs under `tmp/normalized/`.

## Testing
Run the suite (mocks ffmpeg/ffprobe):
```bash
pytest
```

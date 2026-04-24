"""
MP4 → transcript using openai/whisper-large-v3-turbo (HuggingFace, local).
Supports Polish / English / mixed audio. Optimised for Apple M2 via MPS.

Usage:
    python transcribe.py recording.mp4
    python transcribe.py recording.mp4 --format text
    python transcribe.py recording.mp4 --output my_notes.md
"""

import argparse
import os
import subprocess
import sys
import tempfile
from pathlib import Path

import torch
from transformers import pipeline


MODEL_ID = "openai/whisper-large-v3-turbo"
SAMPLE_RATE = 16_000  # Whisper requires 16 kHz mono


def pick_device() -> tuple[str, torch.dtype]:
    if torch.backends.mps.is_available():
        return "mps", torch.float16
    if torch.cuda.is_available():
        return "cuda", torch.float16
    return "cpu", torch.float32


def check_audio_stream(mp4_path: Path) -> bool:
    """Return True if the file contains at least one audio stream."""
    result = subprocess.run(
        ["ffprobe", "-v", "error", "-select_streams", "a",
         "-show_entries", "stream=codec_type", "-of", "csv=p=0", str(mp4_path)],
        capture_output=True, text=True,
    )
    return bool(result.stdout.strip())


def extract_audio(mp4_path: Path, wav_path: Path) -> None:
    """Convert MP4 to 16 kHz mono WAV with ffmpeg."""
    if not check_audio_stream(mp4_path):
        raise RuntimeError(
            f"No audio stream found in '{mp4_path.name}'.\n"
            "The file may be a video-only recording. Check the source file."
        )
    cmd = [
        "ffmpeg", "-y",
        "-i", str(mp4_path),
        "-ar", str(SAMPLE_RATE),
        "-ac", "1",
        "-vn",
        str(wav_path),
    ]
    result = subprocess.run(cmd, capture_output=True)
    if result.returncode != 0:
        print(result.stderr.decode(), file=sys.stderr)
        raise RuntimeError("ffmpeg audio extraction failed.")


def load_pipeline(device: str, dtype: torch.dtype):
    print(f"Loading {MODEL_ID} on {device} …")
    return pipeline(
        "automatic-speech-recognition",
        model=MODEL_ID,
        device=device,
        dtype=dtype,
    )


def transcribe_file(wav_path: Path, pipe) -> str:
    print("Transcribing … (this takes a few minutes for long recordings)")
    result = pipe(
        str(wav_path),
        chunk_length_s=30,        # Whisper context window
        stride_length_s=5,        # overlap between chunks
        batch_size=8,             # parallel chunks; reduce to 4 if RAM is tight
        generate_kwargs={
            "task": "transcribe",
            "language": None,     # auto-detect per chunk → handles mixed PL/EN
        },
        return_timestamps=False,
    )
    return result["text"].strip()


def format_transcript(text: str, source: str, fmt: str) -> str:
    if fmt == "markdown":
        filename = Path(source).name
        return (
            f"# Transcript — {filename}\n\n"
            f"**Source:** `{source}`\n\n"
            "---\n\n"
            + text
        )
    return text  # plain text


def save(content: str, output_path: Path) -> None:
    output_path.write_text(content, encoding="utf-8")
    print(f"Saved → {output_path}")


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Transcribe an MP4 file to text using Whisper large-v3-turbo."
    )
    parser.add_argument("mp4_file", help="Path to the input MP4 file")
    parser.add_argument(
        "--output", "-o",
        help="Output file path (default: same name as MP4 with .md or .txt extension)",
    )
    parser.add_argument(
        "--format", "-f",
        choices=["markdown", "text"],
        default="markdown",
        help="Output format (default: markdown)",
    )
    args = parser.parse_args()

    mp4_path = Path(args.mp4_file).expanduser().resolve()
    if not mp4_path.exists():
        sys.exit(f"File not found: {mp4_path}")

    ext = ".md" if args.format == "markdown" else ".txt"
    output_path = Path(args.output) if args.output else mp4_path.with_suffix(ext)

    device, dtype = pick_device()
    pipe = load_pipeline(device, dtype)

    with tempfile.NamedTemporaryFile(suffix=".wav", delete=False) as tmp:
        wav_path = Path(tmp.name)

    try:
        print(f"Extracting audio from {mp4_path.name} …")
        extract_audio(mp4_path, wav_path)

        transcript = transcribe_file(wav_path, pipe)
        content = format_transcript(transcript, str(mp4_path), args.format)
        save(content, output_path)

    finally:
        wav_path.unlink(missing_ok=True)


if __name__ == "__main__":
    main()

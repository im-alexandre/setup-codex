from __future__ import annotations

import argparse
import os
import shutil
import subprocess
import sys
import tempfile
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable


SUPPORTED_EXTENSIONS = {".ogg", ".mp3", ".wav", ".m4a"}
DEFAULT_MODEL = "medium"


@dataclass
class TranscriptResult:
    source: Path
    text: str | None
    error: str | None = None


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Transcribe local audio files into one Markdown file."
    )
    parser.add_argument(
        "--input",
        default=".",
        help="Path to an audio file or a directory containing audio files.",
    )
    parser.add_argument(
        "--output",
        default="transcricoes.md",
        help="Output Markdown path.",
    )
    parser.add_argument(
        "--language",
        default="pt",
        help="Language hint for Whisper. Default: pt.",
    )
    parser.add_argument(
        "--model",
        default=DEFAULT_MODEL,
        help=f"Whisper model size. Default: {DEFAULT_MODEL}.",
    )
    parser.add_argument(
        "--ffmpeg",
        default=None,
        help="Optional explicit ffmpeg binary path.",
    )
    return parser.parse_args()


def _ffmpeg_works(candidate: str) -> bool:
    try:
        subprocess.run(
            [candidate, "-version"],
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
            check=True,
        )
        return True
    except Exception:  # noqa: BLE001
        return False


def find_ffmpeg(explicit_path: str | None = None) -> str:
    candidates: list[str] = []

    if explicit_path:
        candidates.append(explicit_path)
    if os.environ.get("FFMPEG_BIN"):
        candidates.append(os.environ["FFMPEG_BIN"])

    direct = shutil.which("ffmpeg")
    if direct:
        candidates.append(direct)

    local_appdata = Path.home() / "AppData" / "Local"
    candidates.extend(
        [
            str(local_appdata / "Microsoft" / "WindowsApps" / "ffmpeg.exe"),
            str(
                local_appdata
                / "Microsoft"
                / "WinGet"
                / "Packages"
                / "Gyan.FFmpeg_Microsoft.Winget.Source_8wekyb3d8bbwe"
                / "ffmpeg-8.1-full_build"
                / "bin"
                / "ffmpeg.exe"
            ),
        ]
    )

    for candidate in candidates:
        if candidate and _ffmpeg_works(candidate):
            return candidate

    package_root = (
        local_appdata
        / "Microsoft"
        / "WinGet"
        / "Packages"
        / "Gyan.FFmpeg_Microsoft.Winget.Source_8wekyb3d8bbwe"
    )
    if package_root.is_dir():
        try:
            matches = sorted(package_root.glob("ffmpeg-*/bin/ffmpeg.exe"))
            if matches:
                for match in matches:
                    if _ffmpeg_works(str(match)):
                        return str(match)
        except PermissionError:
            pass

    raise FileNotFoundError(
        "ffmpeg nao encontrado. Instale o FFmpeg ou ajuste o PATH antes de usar esta skill."
    )


def configure_ffmpeg_environment(ffmpeg_bin: str) -> None:
    ffmpeg_dir = str(Path(ffmpeg_bin).resolve().parent)
    current_path = os.environ.get("PATH", "")
    if ffmpeg_dir.lower() not in current_path.lower():
        os.environ["PATH"] = ffmpeg_dir + os.pathsep + current_path


def collect_audio_files(input_path: Path) -> list[Path]:
    if not input_path.exists():
        raise FileNotFoundError(f"Caminho inexistente: {input_path}")

    if input_path.is_file():
        if input_path.suffix.lower() not in SUPPORTED_EXTENSIONS:
            raise ValueError(
                f"Extensao nao suportada: {input_path.suffix}. "
                f"Suportadas: {', '.join(sorted(SUPPORTED_EXTENSIONS))}"
            )
        return [input_path]

    files = sorted(
        path
        for path in input_path.iterdir()
        if path.is_file() and path.suffix.lower() in SUPPORTED_EXTENSIONS
    )
    if not files:
        raise FileNotFoundError(
            f"Nenhum audio suportado encontrado em {input_path}. "
            f"Use arquivos com extensoes: {', '.join(sorted(SUPPORTED_EXTENSIONS))}"
        )
    return files


def convert_to_wav(ffmpeg_bin: str, source: Path, temp_dir: Path) -> Path:
    wav_path = temp_dir / f"{source.stem}.wav"
    cmd = [
        ffmpeg_bin,
        "-y",
        "-loglevel",
        "error",
        "-i",
        str(source),
        "-ac",
        "1",
        "-ar",
        "16000",
        "-c:a",
        "pcm_s16le",
        str(wav_path),
    ]
    subprocess.run(cmd, check=True)
    return wav_path


def load_model(model_name: str):
    import torch
    import whisper

    device = "cuda" if torch.cuda.is_available() else "cpu"
    return whisper.load_model(model_name, device=device), device


def transcribe_files(
    files: Iterable[Path], ffmpeg_bin: str, model_id: str, language: str
) -> list[TranscriptResult]:
    configure_ffmpeg_environment(ffmpeg_bin)
    model, device = load_model(model_id)
    results: list[TranscriptResult] = []

    with tempfile.TemporaryDirectory(prefix="transcribe-local-audio-") as temp_root:
        temp_dir = Path(temp_root)
        for source in files:
            try:
                wav_path = convert_to_wav(ffmpeg_bin, source, temp_dir)
                raw = model.transcribe(
                    str(wav_path),
                    language=language,
                    task="transcribe",
                    fp16=(device == "cuda"),
                    verbose=False,
                )
                text = normalize_text(raw.get("text", ""))
                results.append(TranscriptResult(source=source, text=text))
            except Exception as exc:  # noqa: BLE001
                results.append(TranscriptResult(source=source, text=None, error=str(exc)))

    return results


def normalize_text(text: str) -> str:
    cleaned = text.replace("\r\n", "\n").replace("\r", "\n").strip()
    return cleaned or "[inaudivel]"


def write_markdown(results: list[TranscriptResult], output_path: Path) -> None:
    lines: list[str] = ["# Transcricoes", ""]
    for result in results:
        lines.append(f"## {result.source.stem}")
        lines.append("")
        if result.error:
            lines.append(f"[erro] {result.error}")
        else:
            lines.append(result.text or "[inaudivel]")
        lines.append("")

    output_path.parent.mkdir(parents=True, exist_ok=True)
    output_path.write_text("\n".join(lines).rstrip() + "\n", encoding="utf-8")


def summarize(results: list[TranscriptResult], output_path: Path) -> int:
    success = [result for result in results if not result.error]
    failures = [result for result in results if result.error]

    print(f"Arquivo gerado: {output_path}")
    print(f"Transcritos com sucesso: {len(success)}")
    print(f"Falhas: {len(failures)}")
    for result in success:
        print(f"[ok] {result.source.name}")
    for result in failures:
        print(f"[erro] {result.source.name}: {result.error}")

    return 0 if not failures else 1


def main() -> int:
    args = parse_args()
    input_path = Path(args.input).resolve()
    output_path = Path(args.output).resolve()

    files = collect_audio_files(input_path)
    ffmpeg_bin = find_ffmpeg(args.ffmpeg)
    results = transcribe_files(
        files=files,
        ffmpeg_bin=ffmpeg_bin,
        model_id=args.model,
        language=args.language,
    )
    write_markdown(results, output_path)
    return summarize(results, output_path)


if __name__ == "__main__":
    sys.exit(main())

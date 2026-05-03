---
name: transcribe-local-audio
description: Transcribe local audio files into a consolidated Markdown transcript using local speech-to-text. Use when Codex needs to transcribe `.ogg`, `.mp3`, `.wav`, or `.m4a` files from a single file path or a directory, especially for requests such as converting WhatsApp voice notes to text, transcribing a folder of audio locally, or producing one `.md` with one section per audio.
---

# Transcribe Local Audio

Use this skill to turn local audio files into one Markdown file with a section per audio. The deterministic work lives in `scripts/transcribe_audio.py`; use the script first, then improve readability without changing meaning.

## Workflow

1. Resolve the input path.
If the user gave a file, transcribe only that file.
If the user gave a directory, transcribe every supported audio file in that directory, ordered by name.
If the user gave no path, use the current working directory.

2. Run the transcription script.
Use:

```powershell
python /absolute/path/to/transcribe-local-audio/scripts/transcribe_audio.py --input "<path>" --output "<path-to-output-md>" --language pt
```

Defaults:
- input: current directory
- output: `transcricoes.md` in the current directory
- language: `pt`
- model: `medium` by default

3. Review the generated Markdown.
Improve punctuation, paragraph breaks, and obvious formatting issues only when that materially helps readability.
Preserve wording and meaning.
Do not summarize.
Do not omit repeated phrases that may be intentional.
If audio is unclear, keep markers such as `[inaudivel]` instead of inventing content.

4. Report execution clearly.
State which files were transcribed, where the Markdown was written, and whether any files failed.

## Script Contract

The script:
- accepts one file or one directory
- supports `.ogg`, `.mp3`, `.wav`, `.m4a`
- converts each input to temporary mono 16 kHz WAV with `ffmpeg`
- transcribes with local Whisper via `openai-whisper`
- writes one Markdown file with this structure:

```md
# Transcricoes

## file-name-1

texto...

## file-name-2

texto...
```

- continues after per-file failures and lists them at the end of stdout

## Notes

- Prefer this skill over Word-based transcription when the user wants a local workflow.
- Use Portuguese as the default language unless the user explicitly asks for another language.
- The script is designed to run locally and may download the Whisper model on first use.
- The current default is `medium` on CPU unless the user explicitly overrides `--model`.

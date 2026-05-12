# Video Transcript Extractor

A Claude Code Skill that downloads videos and extracts transcripts from 10+ platforms, generating clean Word documents.

## Supported Platforms

YouTube, Bilibili (B站), Douyin (抖音), Toutiao (头条), Ixigua (西瓜视频), XiaoHongShu (小红书), Zhihu (知乎), WeChat Channels (微信视频号), Kuaishou (快手), Weibo (微博), and any generic video page.

## How It Works

Three-tier transcript extraction:

1. **Captions** — extracts subtitles via yt-dlp (fastest, most accurate)
2. **Whisper STT** — falls back to faster-whisper speech recognition when no captions
3. **Metadata** — last resort, generates document from video title/description

## Install as Claude Code Skill

```bash
# Clone to your Claude Code skills directory
git clone https://github.com/<your-username>/video-transcript-extractor.git \
  ~/.claude/skills/video-transcript-extractor
```

Then in Claude Code, just say: *"Extract the transcript from this video: https://..."*

## Standalone Usage

```bash
# Install dependencies
pip install -r requirements.txt
playwright install chromium  # only needed for WeChat Channels

# System dependency
brew install ffmpeg  # macOS
# sudo apt install ffmpeg  # Linux

# Run
python scripts/extract_transcript.py <VIDEO_URL> [OUTPUT_DIR] [--model base] [--lang zh]
```

### Options

| Flag | Default | Description |
|------|---------|-------------|
| `--model` | `base` | Whisper model: tiny, base, small, medium, large |
| `--lang` | auto | Force language (e.g., `zh`, `en`, `ja`) |

## Output

- `video-title.mp4` — downloaded video
- `video-title-zh.docx` — transcript (Chinese)
- `video-title-en.docx` — transcript (English, if applicable)

## Prerequisites

- Python 3.9+
- ffmpeg (for audio extraction)
- yt-dlp (video download)

All Python dependencies auto-install on first run.

## License

MIT

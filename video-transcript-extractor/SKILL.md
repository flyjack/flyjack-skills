---
name: video-transcript-extractor
description: Use when a user provides a video URL and wants to download the video and extract the transcript as a document. Supports YouTube, Bilibili (B站), Douyin (抖音), Toutiao (头条), Ixigua (西瓜), XiaoHongShu (小红书), Zhihu (知乎), WeChat Channels (微信视频号), Kuaishou (快手), Weibo (微博), and generic video pages. Downloads video as MP4/MP3/WebM/MKV, extracts transcript via captions or Whisper speech-to-text, auto-detects language, and generates formatted documents in DOCX/TXT/PDF/Markdown.
---

# Video Transcript Extractor

## Overview

Extract clean transcripts from videos as formatted Word documents. Supports 10+ major video platforms including Chinese platforms. Automatically detects the video's language, translates to Chinese if needed, and delivers well-formatted .docx files.

**Core principle:** Convert video content into searchable, shareable text with a single command — even when no captions are available (falls back to Whisper AI speech recognition).

## Supported Platforms

| Platform | Method | Notes |
|----------|--------|-------|
| YouTube | yt-dlp | Full support including auto-captions |
| Bilibili (B站) | yt-dlp | May need cookies for some videos |
| Douyin (抖音) | yt-dlp | Short video platform |
| Toutiao (头条) | yt-dlp | News/video platform |
| Ixigua (西瓜视频) | yt-dlp | Long-form video |
| XiaoHongShu (小红书) | yt-dlp | Lifestyle platform |
| Zhihu (知乎) | yt-dlp | Knowledge platform |
| WeChat Channels (微信视频号) | Playwright | Browser automation required |
| Kuaishou (快手) | yt-dlp | Short video platform |
| Weibo (微博) | yt-dlp | Social media |
| Other sites | yt-dlp → Playwright fallback | Auto-detected |

## When to Use

This skill triggers when:
- User provides a video URL from any supported platform and wants the transcript/content
- User wants to download a video and extract its spoken content
- User needs a clean, readable transcript without timestamps
- User has a lecture, talk, presentation, or educational video to document

**Not for:**
- Subtitle/SRT file extraction (this produces clean .docx documents)
- Live streams that are still in progress

## What You'll Deliver

**Video/Audio** (user chooses format):
- MP4 (default), MP3 (audio only), WebM, MKV

**Transcript documents** (user chooses format):
- DOCX (default) — professionally formatted with proper margins, fonts, paragraph spacing
- TXT — plain text with title header
- PDF — formatted PDF with CJK font support
- Markdown — clean `.md` with frontmatter

All formats support multiple languages per video (e.g., original + Chinese translation).

## The Workflow

### Step 1: Prepare
1. Get the video URL from user
2. Detect the platform automatically
3. Check/auto-install dependencies (yt-dlp, python-docx, langdetect, faster-whisper)

### Step 2: Download Video
1. **Primary (yt-dlp):** Use yt-dlp with platform-specific options (referer headers, format selection)
2. **Fallback (Playwright):** For WeChat Channels or if yt-dlp fails — launch headless browser, find `<video>` element or `.mp4` URL in page source, download with proper headers

### Step 3: Extract Transcript (Three Tiers)

**Tier 1 — Captions (fastest, most accurate):**
- Use yt-dlp to download subtitles/auto-captions
- Clean VTT/SRT format (remove timestamps, HTML tags, duplicates)

**Tier 2 — Whisper Speech-to-Text (when no captions):**
- Extract audio with ffmpeg (16kHz mono WAV)
- Run faster-whisper with auto language detection
- Supports models: tiny/base/small/medium/large (default: base)

**Tier 3 — Metadata Only (last resort):**
- Extract title, description, channel, duration from video metadata
- Create structured document noting captions were unavailable

### Step 4: Generate Documents
1. Detect language of transcript
2. Create documents in user-requested formats (docx/txt/pdf/md)
3. DOCX: professional layout with margins, CJK fonts, paragraph indentation, 1.8x line spacing
4. Whisper segments auto-grouped into logical paragraphs (~5 sentences each)
5. If not Chinese, optionally translate to Chinese and create second set of documents
6. Save with convention: `{video-title}-{lang-code}.{ext}`

## Prerequisites & Setup

### Required:
- **yt-dlp** — video download (usually pre-installed as CLI)
- **ffmpeg** — audio extraction for Whisper
- **python-docx** — Word document creation (auto-installed)
- **langdetect** — language detection (auto-installed)

### Auto-installed on demand:
- **faster-whisper** — speech-to-text when no captions
- **playwright** — browser automation for WeChat Channels
- **requests** — HTTP downloads for Playwright path

User should not manually install anything — the skill handles it.

## Common Issues & Solutions

### No captions available
Not a problem — Whisper automatically kicks in as fallback. The resulting transcript is speech-recognized text (may have minor errors with proper nouns).

### yt-dlp fails for a platform
Playwright browser automation takes over automatically. If both fail, a clear error message is shown.

### Bilibili requires login
Some Bilibili videos need cookies. Set the `--cookies-from-browser` option or provide a cookies file.

### WeChat Channels video not found
Some WeChat Channel videos require being logged in. Try opening the URL in a browser first to verify access.

### Whisper model too slow
Default is `base` (good balance). Use `--model tiny` for speed or `--model medium` for accuracy.

## CLI Usage

```bash
python extract_transcript.py <VIDEO_URL> [OUTPUT_DIR]
       [--model base] [--lang zh]
       [--transcript-format docx,txt,pdf,md]
       [--video-format mp4,mp3,webm,mkv]
```

Examples:
```bash
# YouTube video — default (mp4 + docx)
python extract_transcript.py "https://www.youtube.com/watch?v=xxx"

# Bilibili — all transcript formats + audio extraction
python extract_transcript.py "https://www.bilibili.com/video/BVxxx" ./output \
  --transcript-format docx,txt,pdf,md --video-format mp4,mp3

# Toutiao — Chinese, medium model, markdown only
python extract_transcript.py "https://www.toutiao.com/video/xxx" \
  --lang zh --model medium --transcript-format md
```

### Format Options

| Flag | Values | Default |
|------|--------|---------|
| `--transcript-format` | docx, txt, pdf, md (comma-separated) | docx |
| `--video-format` | mp4, mp3, webm, mkv (comma-separated) | mp4 |
| `--model` | tiny, base, small, medium, large | base |
| `--lang` | zh, en, ja, ko, etc. | auto-detect |

## Example Interaction

**User:** "Extract transcript from this Bilibili video: https://www.bilibili.com/video/BV..."

**You:**
```
🎬 Processing: https://www.bilibili.com/video/BV...
📌 Platform: bilibili
📹 Downloading video via yt-dlp...
✅ Video downloaded: video-title.mp4
🔤 Attempting to extract captions via yt-dlp...
⚠️ No captions available. Falling back to Whisper transcription...
🤖 Loading Whisper model (base)...
📝 Transcribing...
🔤 Detected language: zh (probability: 0.99)
✅ Transcription complete: 245 segments, 3892 characters

Done! Your files:
  📹 video-title.mp4
  📄 video-title-zh.docx
  📊 Transcript source: whisper
```

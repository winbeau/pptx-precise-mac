# Changelog

## 1.2.0 — 2026-11-12

- Added a desktop and CLI image mode: compact by default, or exact lossless pixels.
- Compact mode keeps the source resolution, uses JPEG quality 95 with 4:4:4 color,
  and keeps small or low-color graphics lossless.
- Lossless mode uses PNG row prediction. Both modes discard unreferenced PDF objects.
- Added a conservative placement-and-appearance fallback for direct image matches.
- Corrected PowerPoint's rounded PDF page boxes to the exact slide dimensions.
- Reports distinguish source matching, output pixel verification and lossless preservation.

## 1.1.0 — 2026-11-11

First public release.

- Native Swift/AppKit desktop window with status, export logs and batch conversion.
- Custom application icon and self-contained macOS 11+ Apple Silicon package.
- PowerPoint native PDF rendering from isolated copies of saved PPTX files.
- Conservative, per-slide original image restoration with pixel verification.
- Validation of text, dimensions, fonts, page commands and nested vector streams.
- Automatic backups, source-change detection, duplicate-name rejection and process locking.
- Self-contained Apple Silicon app, CLI, source and build script.

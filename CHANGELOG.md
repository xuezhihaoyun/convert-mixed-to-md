# Changelog

## Unreleased

- Added Qwen VL OCR fallback for scanned PDFs when MinerU is unavailable or fails.
- Added direct image-to-Markdown OCR support for common image formats.
- Added optional post-conversion polish and audit flow via `--polish`.
- Added polish profiles: `legal`, `judgment`, `notice`, `book`, and `generic`.
- Added `--report` and `--artifact-level` for audit reports and intermediate artifacts.
- Added preservation checks so polish output is only applied when non-whitespace text remains unchanged.
- Added `anthropic` as an optional dependency for Qwen VL OCR.
- Updated macOS and Windows launch scripts to reflect the new supported formats and dependencies.

# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [1.2.0] - 2025-12-24

### Added
- **Performance optimization**: Buffered paragraph writing reduces Pages automation from O(n²) to O(n) for large documents.
- **Wide table support**: Tables with more than 26 columns now work correctly using Excel-style addressing (AA, AB, etc.).
- **Output directory creation**: Parent directories are created automatically if missing.
- **Safer overwrite**: `--overwrite` now uses backup-and-swap instead of delete-then-move.
- **Concurrency lock**: Uses flock on `/tmp/docx2pages.lock` to prevent parallel runs (default on).
- **`--no-lock` flag**: Disables concurrency lock for advanced users.
- **`--no-wait` flag**: Fail immediately if lock is held instead of waiting.
- **`--batch-size N` flag**: Configure paragraph batch size for Pages writer (default: 50).
- **`--json-summary <path|->` flag**: Outputs machine-readable JSON summary.
  - Use `-` to write JSON to stdout (human output goes to stderr).
  - JSON includes: toolVersion, paths, parse stats, writer results, timing, success/error.
- **Style lookup caching**: Pages writer caches style lookups to avoid repeated AppleScript calls.
- **Per-paragraph error recovery**: Writer continues on individual paragraph failures, tracking error count.
- **Table row chunking**: Tables are processed in 50-row chunks for better reliability.
- **Parser edge case handling**: Graceful handling of empty DOCX, missing styles.xml, malformed XML.
- **Whitespace fidelity**: Parser now preserves tabs (`\t`) and soft line breaks (`\n`).
- **New fixtures**: `large.docx` (300+ paragraphs), `wide_table.docx` (35 columns), `whitespace.docx`, `empty.docx`, `minimal.docx`.
- **Smoke test script**: `scripts/smoke_test.sh` builds and tests all fixtures.
  - `--quick` mode skips large fixtures for faster development testing.
  - `--no-build` mode skips the build step.
  - Per-fixture timing and summary statistics.
- **Pages discovery**: Checks multiple locations for Pages.app (not just /Applications).
- **Performance documentation**: README now includes expected timing and optimization tips.

### Changed
- Pages writer uses buffered writes with batch style application.
- Table cell addressing uses `colIndexToLetters()` for unlimited columns.
- Version bumped to 1.2.0.

### Fixed
- Large documents no longer cause exponential slowdown.
- **True O(n) paragraph writing**: Uses insertionPoints to append without reading entire body.
- Tables with >26 columns no longer fail or produce incorrect output.
- Output to non-existent directories no longer fails silently.
- Overwrite race condition eliminated (no gap between delete and move).
- Parser handles malformed DOCX files gracefully with warnings instead of crashes.
- **Strict mode now fails on**:
  - Paragraph styling errors (errorCount > 0)
  - Severe parser warnings (invalid ZIP, missing document.xml, malformed XML)
  - Zero blocks from non-trivial files

## [1.1.0] - 2025-12-24

### Added
- **Template safety**: Template is copied at filesystem level before Pages automation. Original template is never opened by Pages.
- **Atomic writes**: Output is written to a temp file first, then atomically moved to final location.
- **`--overwrite` flag**: Explicitly opt-in to overwrite existing output files.
- **`--strict` mode**: Fail with non-zero exit on:
  - Style pollution (new styles appearing in output)
  - Table fallback to text
  - List fallback to formatted text (when lists exist in document)
- **`--preserve-breaks` flag**: Convert page/section breaks to blank paragraphs instead of dropping them.
- **Native list styles**: Lists use template paragraph styles (`Bullet`, `Numbered`, etc.) when available.
- **Break tracking**: Reports count of dropped page/section breaks in stats.
- **Table fallback tracking**: Reports when tables fall back to text rendering.
- **Unused style reporting**: Shows which template styles were not used (informational).
- **Symmetric style validation**: Checks both pollution (new styles) and unused styles.
- **Reliable cleanup**: Documents are always closed properly, even on errors.

### Changed
- Pages writer now uses `--doc` and `--json` argument format.
- Baseline styles captured immediately after opening document.
- Better error messages with specific failure reasons.

### Fixed
- Template file is never modified, even if Pages crashes mid-operation.
- Partial output files are cleaned up on failure.

## [1.0.0] - 2025-12-24

### Added
- Initial release
- Convert DOCX to Pages using template styles
- Support for headings (1-9), Title, Subtitle
- Support for bulleted and numbered lists with nesting
- Support for tables
- Heading level saturation (maps deep headings to template max)
- `--prefix-deep-headings` option for visible heading levels
- `--verbose` logging
- Style pollution detection and warning
- Python DOCX parser with heading style inheritance chain walking
- JXA Pages automation
- Test fixtures for all supported element types

### Known Limitations
- Merged table cells not preserved
- Images not supported (dropped silently)
- Inline formatting not preserved (bold, italic, etc.)
- Footnotes/endnotes not supported
- Text boxes and shapes not supported
- Comments and tracked changes not supported
- Headers/footers not transferred

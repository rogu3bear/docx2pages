# docx2pages

A macOS command-line tool that converts Word documents (.docx) to Pages documents (.pages) using a template's native styles—without style pollution.

## Overview

When you open a .docx file directly in Pages, it imports the Word styles, creating duplicates and "imported" variants that pollute your style list. This tool avoids that by:

1. **Parsing** the DOCX file to extract pure structure (headings, paragraphs, lists, tables)
2. **Copying** your template at the filesystem level (template is never opened)
3. **Applying** only the template's existing styles to the content

The result: a clean Pages document that looks like it was created natively in Pages.

## Features

- All heading levels (1-9) with style saturation for templates with fewer levels
- Title and Subtitle styles
- Bulleted lists (native styles when available)
- Numbered lists with nesting (native styles when available)
- Tables with cell content preservation (supports >26 columns)
- Tabs and soft line breaks preserved in text
- Document order preservation
- Custom styles derived from headings (follows style inheritance chain)
- No Word style pollution in output
- Template safety (filesystem copy, never opened by Pages)
- Strict mode for CI/CD validation
- Concurrency lock prevents parallel runs
- Machine-readable JSON summary output

### Not Supported (by design)

- Inline formatting (bold, italic, underline) — dropped silently
- Footnotes and endnotes — dropped silently
- Text boxes and shapes — dropped silently
- Comments and tracked changes — dropped silently
- Images and media — dropped silently
- Section-level layout reconstruction — dropped silently

## Requirements

- macOS 12.0 or later
- Pages.app installed
- Python 3.x (for DOCX parsing)
- Swift 5.9+ toolchain (for building from source only)

## Installation

### Option 1: Download Release (Recommended)

Download the latest release from [GitHub Releases](https://github.com/rogu3bear/docx2pages/releases):

```bash
# Download and unzip
curl -L https://github.com/rogu3bear/docx2pages/releases/latest/download/docx2pages-X.Y.Z-macos.zip -o docx2pages.zip
unzip docx2pages.zip

# Run directly (no Swift toolchain needed)
./docx2pages-X.Y.Z-macos/docx2pages -i doc.docx -o out.pages -t template.pages
```

The release package includes the binary and required scripts. No Swift toolchain needed.

### Option 2: Build from Source

```bash
# Clone the repository
git clone https://github.com/rogu3bear/docx2pages.git
cd docx2pages

# Build the CLI
swift build -c release

# The binary is at .build/release/docx2pages
# Optionally copy to your PATH:
cp .build/release/docx2pages /usr/local/bin/
```

### Grant Automation Permissions

On first run, macOS will prompt you to grant automation permissions for Pages. You can also pre-authorize via:

**System Settings → Privacy & Security → Automation → Terminal (or your terminal app) → Pages**

## Usage

```bash
docx2pages --input <file.docx> --output <file.pages> --template <template.pages>
```

### Options

| Option | Short | Description |
|--------|-------|-------------|
| `--input` | `-i` | Input DOCX file path (required) |
| `--output` | `-o` | Output .pages file path (required) |
| `--template` | `-t` | Pages template file path (required) |
| `--strict` | | Fail on style pollution or fallback behavior |
| `--overwrite` | | Overwrite output file if it exists |
| `--preserve-breaks` | | Convert page/section breaks to blank paragraphs |
| `--prefix-deep-headings` | | Prefix headings beyond template max with "HN:" |
| `--table-style` | | Name of table style to apply from template |
| `--no-lock` | | Disable concurrency lock (advanced) |
| `--no-wait` | | Fail immediately if lock is held (don't wait) |
| `--batch-size` | | Paragraph batch size for Pages writer (default: 50) |
| `--json-summary` | | Write JSON summary to file (use `-` for stdout) |
| `--scripts-dir` | | Directory containing scripts (for dist package) |
| `--verbose` | `-v` | Enable verbose logging |
| `--help` | `-h` | Show help |
| `--version` | | Show version |

### Examples

```bash
# Basic conversion
docx2pages -i report.docx -o report.pages -t ~/Templates/Corporate.pages

# With verbose output
docx2pages -i report.docx -o report.pages -t ~/Templates/Corporate.pages -v

# Strict mode (fails on any style pollution)
docx2pages -i report.docx -o report.pages -t ~/Templates/Corporate.pages --strict

# Preserve page breaks as blank paragraphs
docx2pages -i report.docx -o report.pages -t ~/Templates/Corporate.pages --preserve-breaks

# Prefix deep headings (H7, H8, H9) when template only has Heading 1-6
docx2pages -i thesis.docx -o thesis.pages -t ~/Templates/Academic.pages --prefix-deep-headings

# Output to nested directory (created automatically)
docx2pages -i doc.docx -o /tmp/nested/output/doc.pages -t template.pages

# Generate JSON summary alongside conversion
docx2pages -i doc.docx -o doc.pages -t template.pages --json-summary summary.json

# JSON to stdout (human output to stderr)
docx2pages -i doc.docx -o doc.pages -t template.pages --json-summary - > result.json
```

## Template Requirements

### Required Styles

Your Pages template should include these paragraph styles for best results:

| Style Name | Purpose |
|------------|---------|
| `Body` | Body text (fallback: "Body Text", "Normal") |
| `Title` | Document title |
| `Subtitle` | Document subtitle |
| `Heading` or `Heading 1` | Level 1 headings |
| `Heading 2` ... `Heading 9` | Additional heading levels |

### List Styles (Recommended)

For native list rendering, include these paragraph styles in your template:

| Style Name | Purpose |
|------------|---------|
| `Bullet` | Bulleted list items (alternatives: "Bulleted", "Bulleted List", "Bullets") |
| `Numbered` | Numbered list items (alternatives: "Numbered List", "Numbers") |

**Without list styles**: Lists will be rendered as formatted text with `•` or `1.` prefixes and indentation. Visually correct but not "true" Pages list objects.

**With list styles**: Lists use the native paragraph styles from your template, giving you full control over list appearance.

### Creating a Template

1. Open Pages and create a new blank document
2. Go to **Format → Paragraph Styles**
3. Create or modify styles: Body, Title, Heading, Heading 2, etc.
4. For lists: Create "Bullet" and "Numbered" paragraph styles
5. Save as your template file

## How It Works

### Architecture

```
┌─────────────┐      ┌──────────────┐      ┌─────────────────┐
│  DOCX File  │─────▶│ Python Parser │─────▶│  JSON Blocks    │
└─────────────┘      │ (parse_docx) │      │  (intermediate) │
                     └──────────────┘      └────────┬────────┘
                                                    │
┌─────────────┐      ┌──────────────┐               │
│  Template   │─────▶│  Filesystem  │               │
│   (.pages)  │      │     Copy     │               │
└─────────────┘      └──────┬───────┘               │
                            │                       │
                     ┌──────▼───────┐               │
                     │ Pages Writer │◀──────────────┘
                     │    (JXA)     │
                     └──────┬───────┘
                            │
                     ┌──────▼───────┐
                     │ Output File  │
                     │   (.pages)   │
                     └──────────────┘
```

**Key safety feature**: The template is copied at the filesystem level before any Pages automation. Pages only opens the copy, never the original template.

### Block Types

The parser emits these block types:

```json
{"type": "title", "text": "Document Title"}
{"type": "subtitle", "text": "Subtitle Text"}
{"type": "heading", "level": 1, "text": "Chapter One"}
{"type": "paragraph", "text": "Body text..."}
{"type": "list", "ordered": false, "items": [{"text": "Item", "level": 0}]}
{"type": "table", "rows": [["Cell 1", "Cell 2"], ["Cell 3", "Cell 4"]]}
{"type": "break"}
```

### Style Mapping

The tool maps Word styles to Pages template styles:

| Word Style | Pages Style (preference order) |
|------------|-------------------------------|
| Normal / Body | Body → Body Text → Normal |
| Title | Title → Heading 1 |
| Subtitle | Subtitle → Body |
| Heading 1 | Heading → Heading 1 |
| Heading N (2-9) | Heading N (saturates at template max) |
| Bulleted list | Bullet → Bulleted → Bulleted List → Bullets (fallback: text) |
| Numbered list | Numbered → Numbered List → Numbers (fallback: text) |

**Saturation**: If your template only has Heading 1-3 and the document has Heading 5, it maps to Heading 3.

## Strict Mode

Use `--strict` to enforce clean conversions:

```bash
docx2pages -i doc.docx -o out.pages -t template.pages --strict
```

In strict mode, the tool will **fail with a non-zero exit code** if:
- Any new paragraph styles appear in the output that weren't in the template
- Any table falls back to text rendering
- Any list falls back to text rendering (when template lacks list styles)

This is useful for CI/CD pipelines to catch unexpected style pollution.

Without `--strict`, fallbacks are reported as warnings but the conversion succeeds.

## Concurrency

By default, docx2pages acquires an exclusive lock on `/tmp/docx2pages.lock` to prevent multiple simultaneous conversions. This avoids conflicts when Pages is processing documents.

Use `--no-lock` to disable this behavior if you need to run multiple conversions in controlled environments.

## JSON Summary

Use `--json-summary <path>` to output a machine-readable summary:

```bash
docx2pages -i doc.docx -o out.pages -t template.pages --json-summary result.json
```

Or use `-` to write JSON to stdout (human output goes to stderr):

```bash
docx2pages -i doc.docx -o out.pages -t template.pages --json-summary - 2>/dev/null
```

The JSON includes:
- `toolVersion`: Version string
- `input`, `output`, `template`: File paths
- `strict`: Whether strict mode was enabled
- `parseStats`: Headings, paragraphs, lists, tables counts
- `writeResult`: Styles used, pollution detected, warnings
- `elapsedSeconds`: Conversion time
- `success`: Boolean
- `error`: Error message if failed

## Test Fixtures

The `fixtures/` directory contains test DOCX files:

| File | Description |
|------|-------------|
| `all_headings.docx` | Title, Subtitle, Heading 1-9, custom derived heading |
| `mixed_lists.docx` | Bulleted and numbered lists with nesting |
| `tables.docx` | Multiple tables (3x3 and 5x4) |
| `comprehensive.docx` | All element types combined |
| `large.docx` | 300+ paragraphs for performance testing |
| `wide_table.docx` | Table with 35 columns (tests >26 column addressing) |
| `whitespace.docx` | Tabs and soft line breaks |
| `empty.docx` | Empty document (edge case testing) |
| `minimal.docx` | Single paragraph document (edge case testing) |

### Generate Fixtures

```bash
python3 scripts/create_fixtures.py
```

### Smoke Test

Run the smoke test to verify all fixtures convert successfully:

```bash
# Requires a Pages template file
TEMPLATE=/path/to/template.pages scripts/smoke_test.sh
```

The smoke test:
1. Builds the release binary
2. Regenerates all fixtures
3. Converts each fixture with `--strict --overwrite`
4. Reports pass/fail for each

## CI/CD Integration

### GitHub Actions Workflows

The repository includes two CI workflows:

**`ci.yml`** - Runs on every push and PR:
- Parser tests on Ubuntu and macOS (no Pages required)
- Golden test comparison for parser output stability
- Swift build verification on macOS
- CLI surface contract checks

**`pages-integration.yml`** - Manual trigger for full testing:
- Requires a self-hosted macOS runner with Pages installed
- Runs full smoke test with `--strict` mode
- Uploads test outputs as artifacts

### Running Tests Locally

```bash
# Parser golden tests (no Pages required, runs on any platform)
python3 scripts/test_parser_golden.py

# Update golden files after intentional parser changes
python3 scripts/test_parser_golden.py --update

# Build and CLI checks
swift build -c release
.build/release/docx2pages --version
.build/release/docx2pages --help
```

### Integration Tests (Mac with Pages)

Full integration tests require:
- macOS with Pages.app installed
- Automation permission granted
- A Pages template file

```bash
# Run smoke test (requires template)
TEMPLATE=/path/to/template.pages scripts/smoke_test.sh

# Quick mode (skips large fixtures)
TEMPLATE=/path/to/template.pages scripts/smoke_test.sh --quick

# Individual conversion with strict mode
.build/release/docx2pages \
    -i fixtures/comprehensive.docx \
    -o /tmp/test.pages \
    -t /path/to/template.pages \
    --strict
echo $?  # 0 = success
```

## Known Limitations

1. **Merged Cells**: Table cell merging is not preserved; merged cells appear as separate cells.

2. **Images**: Not supported. Images in DOCX are silently dropped.

3. **Headers/Footers**: Not extracted or transferred.

4. **Page Breaks**: By default, page/section breaks are dropped. Use `--preserve-breaks` to convert them to blank paragraphs.

5. **Inline Formatting**: Bold, italic, underline, etc. are not preserved. Only paragraph-level styles are applied.

6. **Nested List Levels**: Nesting is represented by indentation only. True multi-level list numbering (1.1, 1.2, etc.) depends on template list style configuration.

7. **Footnotes/Endnotes**: Not supported; silently dropped.

8. **Text Boxes/Shapes**: Not supported; silently dropped.

9. **Comments/Track Changes**: Not supported; silently dropped.

## Performance

### Expected Timing

| Document Size | Typical Time |
|---------------|--------------|
| Small (1-50 paragraphs) | 2-5 seconds |
| Medium (50-200 paragraphs) | 5-15 seconds |
| Large (200-500 paragraphs) | 15-45 seconds |
| Very large (500+ paragraphs) | 45+ seconds |

Times are dominated by Pages automation overhead, not parsing.

### Optimization Tips

1. **Batch size tuning**: Use `--batch-size N` to adjust paragraph flushing (default: 50). Larger batches may improve throughput for very large documents, but too large may cause memory issues.

2. **Quick smoke testing**: Use `scripts/smoke_test.sh --quick` to skip large fixtures during development.

3. **Non-blocking lock**: Use `--no-wait` to fail immediately if another conversion is running, instead of waiting.

4. **Parallel conversions**: Not recommended due to Pages automation constraints. The default concurrency lock prevents this.

### Scalability Notes

- The tool uses O(n) paragraph writing with buffered flushes
- Style lookups are cached to avoid repeated AppleScript calls
- Tables are processed in chunks (50 rows at a time)
- Memory usage is proportional to document size

## Troubleshooting

### "Pages.app not found"

Install Pages from the Mac App Store. The tool checks:
- `/Applications/Pages.app`
- `/System/Applications/Pages.app`
- `~/Applications/Pages.app`

### "Automation permission denied"

Grant Terminal (or your terminal app) permission to control Pages:
- System Settings → Privacy & Security → Automation → [Your Terminal] → Pages ✓

If the prompt doesn't appear:
1. Open System Settings → Privacy & Security → Automation
2. Click the + button
3. Navigate to your terminal app
4. Check "Pages" in the list

### "Could not find parse_docx.py"

The tool searches for scripts in this order:
1. `--scripts-dir <path>` if specified
2. `<executable>/../scripts/` (dist package layout)
3. `./scripts/` (current working directory)
4. `~/.docx2pages/scripts/`
5. `/usr/local/share/docx2pages/scripts/`

If using the dist package, the wrapper script sets `--scripts-dir` automatically.

### Styles not applying correctly

1. Verify your template has the expected style names (Body, Heading 1, etc.)
2. Run with `-v` to see which styles are detected and mapped
3. Check that the template file is a valid .pages file (not .pages.zip)

### Lists not using native styles

1. Add paragraph styles named "Bullet" and "Numbered" to your template
2. Alternatively: "Bulleted List" / "Numbered List" or similar
3. Run with `-v` to see which list styles were detected

### "Could not acquire lock" / "Another docx2pages process is running"

Another conversion is in progress. Options:

1. **Wait**: By default, the tool waits for the lock to be released
2. **Fail fast**: Use `--no-wait` to fail immediately instead of waiting
3. **Skip lock**: Use `--no-lock` to disable locking entirely (not recommended for parallel use)

## File Structure

```
docx2pages/
├── Package.swift              # Swift package manifest
├── README.md                  # This file
├── CHANGELOG.md               # Version history
├── CONTRIBUTING.md            # Contribution guidelines
├── SECURITY.md                # Security policy
├── LICENSE                    # MIT License
├── .github/
│   ├── workflows/
│   │   ├── ci.yml             # Main CI (parser + build)
│   │   └── pages-integration.yml  # Pages integration tests
│   ├── ISSUE_TEMPLATE/
│   │   ├── bug_report.md
│   │   └── feature_request.md
│   └── PULL_REQUEST_TEMPLATE.md
├── Sources/
│   └── docx2pages/
│       └── main.swift         # CLI entry point and orchestration
├── scripts/
│   ├── parse_docx.py          # DOCX parsing (Python)
│   ├── pages_writer.js        # Pages automation (JXA)
│   ├── create_fixtures.py     # Test fixture generator
│   ├── test_parser_golden.py  # Golden test runner
│   ├── smoke_test.sh          # Integration smoke test
│   ├── package_dist.sh        # Distribution packaging
│   ├── bump_version.sh        # Version bump helper
│   └── release_checklist.md   # Release process guide
└── fixtures/
    ├── *.docx                 # Test DOCX files
    └── golden/                # Parser golden outputs
        └── *.json
```

## Author

James KC Auchterlonie
AuchShop LLC | MLNavigator Inc.

## License

MIT License. See [LICENSE](LICENSE) file.

## Contributing

1. Fork the repository
2. Create a feature branch
3. Add tests for new functionality
4. Submit a pull request

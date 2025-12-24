# Release Checklist

Follow these steps to create a new release of docx2pages.

## Pre-Release

### 1. Version Bump

Update the version in `Sources/docx2pages/main.swift`:

```swift
version: "X.Y.Z"
```

Or use the helper script:

```bash
scripts/bump_version.sh X.Y.Z
```

### 2. Update CHANGELOG

Add a new section to `CHANGELOG.md`:

```markdown
## [X.Y.Z] - YYYY-MM-DD

### Added
- ...

### Changed
- ...

### Fixed
- ...
```

### 3. Run Tests

```bash
# Parser golden tests (no Pages required)
python3 scripts/test_parser_golden.py

# Build
swift build -c release

# CLI sanity checks
.build/release/docx2pages --version
.build/release/docx2pages --help

# Full smoke test (requires Pages + template)
TEMPLATE=/path/to/template.pages scripts/smoke_test.sh
```

### 4. Build Distribution Package

```bash
scripts/package_dist.sh
```

Verify the output:
- `dist/docx2pages-X.Y.Z-macos.zip` exists
- Unzip and test on a clean machine if possible

## Release

### 5. Commit and Tag

```bash
git add -A
git commit -m "Release vX.Y.Z"
git tag -a vX.Y.Z -m "Release vX.Y.Z"
git push origin main --tags
```

### 6. Create GitHub Release

1. Go to GitHub → Releases → "Draft a new release"
2. Select the tag `vX.Y.Z`
3. Title: `docx2pages vX.Y.Z`
4. Description: Copy from CHANGELOG.md
5. Attach: `dist/docx2pages-X.Y.Z-macos.zip`
6. Publish

## Post-Release

### 7. Verify

- [ ] GitHub release page shows correct version
- [ ] Download link works
- [ ] CI passes on the tagged commit

### 8. Announce

- Update any external documentation
- Notify users if breaking changes

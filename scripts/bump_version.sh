#!/bin/bash
#
# Bump the version number in docx2pages.
#
# Usage:
#   scripts/bump_version.sh <new_version>
#
# Example:
#   scripts/bump_version.sh 1.3.0
#
set -e

if [[ -z "$1" ]]; then
    echo "Usage: $0 <new_version>"
    echo "Example: $0 1.3.0"
    exit 1
fi

NEW_VERSION="$1"

# Validate version format
if ! [[ "$NEW_VERSION" =~ ^[0-9]+\.[0-9]+\.[0-9]+$ ]]; then
    echo "ERROR: Invalid version format. Use X.Y.Z (e.g., 1.3.0)"
    exit 1
fi

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
PROJECT_DIR="$(dirname "$SCRIPT_DIR")"
MAIN_SWIFT="$PROJECT_DIR/Sources/docx2pages/main.swift"

# Get current version
CURRENT_VERSION=$(grep -E 'version:\s*"' "$MAIN_SWIFT" | head -1 | sed 's/.*version: *"\([^"]*\)".*/\1/')

if [[ -z "$CURRENT_VERSION" ]]; then
    echo "ERROR: Could not find current version in $MAIN_SWIFT"
    exit 1
fi

echo "Current version: $CURRENT_VERSION"
echo "New version:     $NEW_VERSION"
echo ""

# Confirm
read -p "Proceed with version bump? [y/N] " -n 1 -r
echo ""

if [[ ! $REPLY =~ ^[Yy]$ ]]; then
    echo "Aborted."
    exit 0
fi

# Update main.swift
sed -i '' "s/version: \"$CURRENT_VERSION\"/version: \"$NEW_VERSION\"/" "$MAIN_SWIFT"

# Verify
UPDATED_VERSION=$(grep -E 'version:\s*"' "$MAIN_SWIFT" | head -1 | sed 's/.*version: *"\([^"]*\)".*/\1/')

if [[ "$UPDATED_VERSION" == "$NEW_VERSION" ]]; then
    echo ""
    echo "Version updated successfully!"
    echo ""
    echo "Next steps:"
    echo "  1. Update CHANGELOG.md with release notes"
    echo "  2. Run tests: python3 scripts/test_parser_golden.py"
    echo "  3. Build: swift build -c release"
    echo "  4. Verify: .build/release/docx2pages --version"
    echo "  5. Commit: git commit -am 'Bump version to $NEW_VERSION'"
else
    echo "ERROR: Version update failed"
    exit 1
fi

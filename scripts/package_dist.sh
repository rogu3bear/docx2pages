#!/bin/bash
#
# Package docx2pages for distribution.
#
# Creates a standalone distribution zip that can be used without Swift toolchain.
#
# Usage:
#   scripts/package_dist.sh
#
# Output:
#   dist/docx2pages-<version>-macos.zip
#
set -e

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
PROJECT_DIR="$(dirname "$SCRIPT_DIR")"

cd "$PROJECT_DIR"

# Get version from Swift source
VERSION=$(grep -E 'version:\s*"' Sources/docx2pages/main.swift | head -1 | sed 's/.*version: *"\([^"]*\)".*/\1/')

if [[ -z "$VERSION" ]]; then
    echo "ERROR: Could not extract version from main.swift"
    exit 1
fi

DIST_NAME="docx2pages-${VERSION}-macos"
DIST_DIR="dist/${DIST_NAME}"
DIST_ZIP="dist/${DIST_NAME}.zip"

echo "=========================================="
echo "Packaging docx2pages v${VERSION}"
echo "=========================================="
echo ""

# Step 1: Build release
echo "Step 1: Building release..."
swift build -c release

BINARY="$PROJECT_DIR/.build/release/docx2pages"
if [[ ! -x "$BINARY" ]]; then
    echo "ERROR: Binary not found at $BINARY"
    exit 1
fi
echo "Binary: $BINARY"
echo ""

# Step 2: Create dist directory structure
echo "Step 2: Creating distribution directory..."
rm -rf "$DIST_DIR"
mkdir -p "$DIST_DIR/bin"
mkdir -p "$DIST_DIR/scripts"

# Step 3: Copy files
echo "Step 3: Copying files..."

# Binary
cp "$BINARY" "$DIST_DIR/bin/"
echo "  bin/docx2pages"

# Scripts (required)
cp scripts/parse_docx.py "$DIST_DIR/scripts/"
echo "  scripts/parse_docx.py"

cp scripts/pages_writer.js "$DIST_DIR/scripts/"
echo "  scripts/pages_writer.js"

# Documentation
cp README.md "$DIST_DIR/"
echo "  README.md"

if [[ -f LICENSE ]]; then
    cp LICENSE "$DIST_DIR/"
    echo "  LICENSE"
fi

if [[ -f CHANGELOG.md ]]; then
    cp CHANGELOG.md "$DIST_DIR/"
    echo "  CHANGELOG.md"
fi

# Optional: Include a sample template if available
TEMPLATE_LOCATIONS=(
    "$HOME/Documents/PagesDefaultStyles.pages"
    "$PROJECT_DIR/templates/PagesDefaultStyles.pages"
)

for t in "${TEMPLATE_LOCATIONS[@]}"; do
    if [[ -e "$t" ]]; then
        mkdir -p "$DIST_DIR/templates"
        cp -r "$t" "$DIST_DIR/templates/"
        echo "  templates/$(basename "$t")"
        break
    fi
done

echo ""

# Step 4: Create wrapper script for easier invocation
echo "Step 4: Creating wrapper script..."
cat > "$DIST_DIR/docx2pages" << 'EOF'
#!/bin/bash
# Wrapper script for docx2pages
# Automatically sets --scripts-dir to the bundled scripts
SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
exec "$SCRIPT_DIR/bin/docx2pages" --scripts-dir "$SCRIPT_DIR/scripts" "$@"
EOF
chmod +x "$DIST_DIR/docx2pages"
echo "  docx2pages (wrapper)"
echo ""

# Step 5: Create zip
echo "Step 5: Creating zip archive..."
cd dist
rm -f "${DIST_NAME}.zip"
zip -r "${DIST_NAME}.zip" "${DIST_NAME}" -x "*.DS_Store"
cd "$PROJECT_DIR"

# Verify
if [[ -f "$DIST_ZIP" ]]; then
    SIZE=$(ls -lh "$DIST_ZIP" | awk '{print $5}')
    echo ""
    echo "=========================================="
    echo "Package created successfully!"
    echo "=========================================="
    echo ""
    echo "Output: $DIST_ZIP"
    echo "Size:   $SIZE"
    echo ""
    echo "Contents:"
    unzip -l "$DIST_ZIP" | tail -n +4 | head -n -2
    echo ""
    echo "To use:"
    echo "  1. Unzip: unzip ${DIST_NAME}.zip"
    echo "  2. Run:   ./${DIST_NAME}/docx2pages -i doc.docx -o out.pages -t template.pages"
else
    echo "ERROR: Failed to create zip"
    exit 1
fi

#!/bin/bash
#
# Smoke test for docx2pages
# Builds release and runs conversion for all fixtures.
# Returns non-zero on any failure.
#
# Options:
#   --quick    Skip large fixtures (large.docx) for faster testing
#   --no-build Skip the build step (use existing binary)
#
set -e

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
PROJECT_DIR="$(dirname "$SCRIPT_DIR")"
FIXTURES_DIR="$PROJECT_DIR/fixtures"
OUTPUT_DIR="/tmp/docx2pages_smoke_test"

# Parse arguments
QUICK_MODE=false
SKIP_BUILD=false
for arg in "$@"; do
    case $arg in
        --quick)
            QUICK_MODE=true
            shift
            ;;
        --no-build)
            SKIP_BUILD=true
            shift
            ;;
    esac
done

# Colors for output
RED='\033[0;31m'
GREEN='\033[0;32m'
YELLOW='\033[1;33m'
CYAN='\033[0;36m'
NC='\033[0m' # No Color

echo "=========================================="
echo "docx2pages Smoke Test"
echo "=========================================="
if [[ "$QUICK_MODE" == "true" ]]; then
    echo -e "${YELLOW}Quick mode: skipping large fixtures${NC}"
fi
echo ""

# Check for template file
TEMPLATE="${TEMPLATE:-$HOME/Documents/PagesDefaultStyles.pages}"
if [[ ! -e "$TEMPLATE" ]]; then
    # Try common locations
    for t in \
        "$HOME/Documents/PagesDefaultStyles.pages" \
        "$HOME/Templates/PagesDefaultStyles.pages" \
        "$PROJECT_DIR/templates/PagesDefaultStyles.pages" \
        "/tmp/PagesDefaultStyles.pages"
    do
        if [[ -e "$t" ]]; then
            TEMPLATE="$t"
            break
        fi
    done
fi

if [[ ! -e "$TEMPLATE" ]]; then
    echo -e "${YELLOW}Warning: No template found.${NC}"
    echo "Please set TEMPLATE environment variable or place a .pages template at:"
    echo "  $HOME/Documents/PagesDefaultStyles.pages"
    echo ""
    echo "To create a minimal template:"
    echo "  1. Open Pages and create a new blank document"
    echo "  2. Save it as PagesDefaultStyles.pages"
    echo ""
    echo "Skipping smoke test (template required)."
    exit 0
fi

echo "Template: $TEMPLATE"
echo ""

# Step 1: Build release
if [[ "$SKIP_BUILD" == "true" ]]; then
    echo "Step 1: Skipping build (--no-build)"
else
    echo "Step 1: Building release..."
    cd "$PROJECT_DIR"
    if swift build -c release 2>&1; then
        echo -e "${GREEN}✓ Build succeeded${NC}"
    else
        echo -e "${RED}✗ Build failed${NC}"
        exit 1
    fi
fi

BINARY="$PROJECT_DIR/.build/release/docx2pages"
if [[ ! -x "$BINARY" ]]; then
    echo -e "${RED}✗ Binary not found: $BINARY${NC}"
    exit 1
fi

echo ""

# Step 2: Regenerate fixtures
echo "Step 2: Regenerating fixtures..."
python3 "$SCRIPT_DIR/create_fixtures.py"
echo ""

# Step 3: Create output directory
echo "Step 3: Preparing output directory..."
rm -rf "$OUTPUT_DIR"
mkdir -p "$OUTPUT_DIR"
echo "Output: $OUTPUT_DIR"
echo ""

# Step 4: Run conversions
echo "Step 4: Running conversions..."
echo ""

PASSED=0
FAILED=0
SKIPPED=0
FIXTURES=()
TIMINGS=()
TOTAL_TIME=0

# Fixtures to skip in quick mode
LARGE_FIXTURES=("large")

for docx in "$FIXTURES_DIR"/*.docx; do
    if [[ ! -f "$docx" ]]; then
        continue
    fi

    name=$(basename "$docx" .docx)
    output="$OUTPUT_DIR/${name}.pages"
    FIXTURES+=("$name")

    # Skip large fixtures in quick mode
    if [[ "$QUICK_MODE" == "true" ]]; then
        skip=false
        for large in "${LARGE_FIXTURES[@]}"; do
            if [[ "$name" == "$large" ]]; then
                skip=true
                break
            fi
        done
        if [[ "$skip" == "true" ]]; then
            echo -e "  Skipping $name.docx... ${YELLOW}SKIP${NC} (quick mode)"
            ((SKIPPED++))
            TIMINGS+=("0.00")
            continue
        fi
    fi

    echo -n "  Converting $name.docx... "

    # Capture start time
    start_time=$(python3 -c 'import time; print(time.time())')

    # Run conversion with --strict --overwrite
    if "$BINARY" \
        -i "$docx" \
        -o "$output" \
        -t "$TEMPLATE" \
        --strict \
        --overwrite \
        2>&1 | tail -1; then

        # Capture end time
        end_time=$(python3 -c 'import time; print(time.time())')
        elapsed=$(python3 -c "print(f'{$end_time - $start_time:.2f}')")
        TIMINGS+=("$elapsed")
        TOTAL_TIME=$(python3 -c "print(f'{$TOTAL_TIME + $end_time - $start_time:.2f}')")

        # Check output exists
        if [[ -e "$output" ]]; then
            echo -e "${GREEN}✓ PASS${NC} ${CYAN}(${elapsed}s)${NC}"
            ((PASSED++))
        else
            echo -e "${RED}✗ FAIL (no output)${NC}"
            ((FAILED++))
        fi
    else
        # Capture end time for failed runs too
        end_time=$(python3 -c 'import time; print(time.time())')
        elapsed=$(python3 -c "print(f'{$end_time - $start_time:.2f}')")
        TIMINGS+=("$elapsed")
        TOTAL_TIME=$(python3 -c "print(f'{$TOTAL_TIME + $end_time - $start_time:.2f}')")

        echo -e "${RED}✗ FAIL (error)${NC} ${CYAN}(${elapsed}s)${NC}"
        ((FAILED++))
    fi
done

echo ""
echo "=========================================="
echo "Results"
echo "=========================================="
echo ""
echo "Total fixtures: ${#FIXTURES[@]}"
echo -e "Passed:  ${GREEN}$PASSED${NC}"
echo -e "Failed:  ${RED}$FAILED${NC}"
if [[ $SKIPPED -gt 0 ]]; then
    echo -e "Skipped: ${YELLOW}$SKIPPED${NC}"
fi
echo ""
echo -e "Total time: ${CYAN}${TOTAL_TIME}s${NC}"
echo ""

# Print timing summary for non-zero times
echo "Timing breakdown:"
for i in "${!FIXTURES[@]}"; do
    if [[ "${TIMINGS[$i]}" != "0.00" ]]; then
        printf "  %-20s %ss\n" "${FIXTURES[$i]}" "${TIMINGS[$i]}"
    fi
done
echo ""

if [[ $FAILED -gt 0 ]]; then
    echo -e "${RED}✗ SMOKE TEST FAILED${NC}"
    exit 1
else
    echo -e "${GREEN}✓ SMOKE TEST PASSED${NC}"
    echo ""
    echo "Output files in: $OUTPUT_DIR"
    exit 0
fi

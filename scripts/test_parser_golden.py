#!/usr/bin/env python3
"""
Golden test runner for parse_docx.py.

Compares parser output against known-good golden files to catch regressions.

Usage:
    python3 scripts/test_parser_golden.py          # Run tests
    python3 scripts/test_parser_golden.py --update # Update golden files
"""

import json
import subprocess
import sys
import argparse
from pathlib import Path

SCRIPT_DIR = Path(__file__).parent
PROJECT_DIR = SCRIPT_DIR.parent
FIXTURES_DIR = PROJECT_DIR / "fixtures"
GOLDEN_DIR = FIXTURES_DIR / "golden"
PARSER_SCRIPT = SCRIPT_DIR / "parse_docx.py"


def normalize_json(data):
    """
    Normalize JSON for deterministic comparison.
    - Sort warnings list
    - Ensure consistent key ordering via json.dumps with sort_keys
    """
    if isinstance(data, dict):
        # Recursively normalize nested dicts
        normalized = {}
        for key, value in data.items():
            normalized[key] = normalize_json(value)

        # Sort warnings list if present
        if 'stats' in normalized and 'warnings' in normalized['stats']:
            normalized['stats']['warnings'] = sorted(normalized['stats']['warnings'])

        return normalized
    elif isinstance(data, list):
        return [normalize_json(item) for item in data]
    else:
        return data


def parse_fixture(fixture_path):
    """Run parse_docx.py on a fixture and return parsed JSON."""
    result = subprocess.run(
        [sys.executable, str(PARSER_SCRIPT), str(fixture_path)],
        capture_output=True,
        text=True
    )

    if result.returncode != 0:
        raise RuntimeError(f"Parser failed on {fixture_path}: {result.stderr}")

    return json.loads(result.stdout)


def get_fixtures():
    """Get all fixture DOCX files."""
    return sorted(FIXTURES_DIR.glob("*.docx"))


def run_tests():
    """Run golden tests, return True if all pass."""
    fixtures = get_fixtures()

    if not fixtures:
        print("ERROR: No fixtures found in", FIXTURES_DIR)
        return False

    passed = 0
    failed = 0
    missing = 0

    print(f"Running golden tests for {len(fixtures)} fixtures...")
    print()

    for fixture in fixtures:
        name = fixture.stem
        golden_path = GOLDEN_DIR / f"{name}.json"

        print(f"  {name}...", end=" ")

        if not golden_path.exists():
            print("MISSING (run with --update)")
            missing += 1
            continue

        try:
            # Parse fixture
            actual = normalize_json(parse_fixture(fixture))

            # Load golden
            with open(golden_path) as f:
                expected = normalize_json(json.load(f))

            # Compare
            actual_str = json.dumps(actual, sort_keys=True, indent=2)
            expected_str = json.dumps(expected, sort_keys=True, indent=2)

            if actual_str == expected_str:
                print("OK")
                passed += 1
            else:
                print("FAIL")
                failed += 1

                # Show diff summary
                print(f"    Golden: {golden_path}")
                print(f"    Actual differs. Run with --update to accept changes.")

                # Show brief stats comparison
                if 'stats' in actual and 'stats' in expected:
                    a_stats = actual['stats']
                    e_stats = expected['stats']

                    for key in set(a_stats.keys()) | set(e_stats.keys()):
                        a_val = a_stats.get(key)
                        e_val = e_stats.get(key)
                        if a_val != e_val:
                            print(f"    stats.{key}: expected {e_val}, got {a_val}")

        except Exception as e:
            print(f"ERROR: {e}")
            failed += 1

    print()
    print(f"Results: {passed} passed, {failed} failed, {missing} missing")

    return failed == 0 and missing == 0


def update_goldens():
    """Update all golden files from current parser output."""
    fixtures = get_fixtures()

    if not fixtures:
        print("ERROR: No fixtures found in", FIXTURES_DIR)
        return False

    # Ensure golden directory exists
    GOLDEN_DIR.mkdir(parents=True, exist_ok=True)

    print(f"Updating golden files for {len(fixtures)} fixtures...")
    print()

    for fixture in fixtures:
        name = fixture.stem
        golden_path = GOLDEN_DIR / f"{name}.json"

        print(f"  {name}...", end=" ")

        try:
            # Parse fixture
            data = normalize_json(parse_fixture(fixture))

            # Write golden with consistent formatting
            with open(golden_path, 'w') as f:
                json.dump(data, f, sort_keys=True, indent=2, ensure_ascii=False)
                f.write('\n')  # Trailing newline

            print("updated")

        except Exception as e:
            print(f"ERROR: {e}")
            return False

    print()
    print("Golden files updated successfully.")
    return True


def main():
    parser = argparse.ArgumentParser(description="Golden test runner for parse_docx.py")
    parser.add_argument(
        "--update",
        action="store_true",
        help="Update golden files instead of testing"
    )

    args = parser.parse_args()

    if args.update:
        success = update_goldens()
    else:
        success = run_tests()

    sys.exit(0 if success else 1)


if __name__ == "__main__":
    main()

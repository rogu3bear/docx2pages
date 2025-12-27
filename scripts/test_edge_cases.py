#!/usr/bin/env python3
"""
Edge case tests for docx2pages security and robustness fixes.

Tests:
1. XXE protection (defusedxml blocks external entity attacks)
2. Control character sanitization
3. Parser error handling

Usage:
    python3 scripts/test_edge_cases.py
"""

import sys
import tempfile
import zipfile
from pathlib import Path

# Add scripts directory to path for imports
SCRIPT_DIR = Path(__file__).parent
sys.path.insert(0, str(SCRIPT_DIR))

from parse_docx import sanitize_text

# Test counters
passed = 0
failed = 0


def test(name, condition, error_msg=""):
    """Simple test helper."""
    global passed, failed
    if condition:
        print(f"  {name}... OK")
        passed += 1
    else:
        print(f"  {name}... FAIL")
        if error_msg:
            print(f"    {error_msg}")
        failed += 1


def test_sanitize_text():
    """Test control character sanitization."""
    print("\nTesting control character sanitization...")

    # Null bytes should be removed
    test(
        "Null byte removal",
        sanitize_text("Hello\x00World") == "HelloWorld"
    )

    # Bell character should be removed
    test(
        "Bell character removal",
        sanitize_text("Hello\x07World") == "HelloWorld"
    )

    # Backspace should be removed
    test(
        "Backspace removal",
        sanitize_text("Hello\x08World") == "HelloWorld"
    )

    # Vertical tab should be removed
    test(
        "Vertical tab removal",
        sanitize_text("Hello\x0BWorld") == "HelloWorld"
    )

    # Form feed should be removed
    test(
        "Form feed removal",
        sanitize_text("Hello\x0CWorld") == "HelloWorld"
    )

    # Tab should be preserved
    test(
        "Tab preservation",
        sanitize_text("Hello\tWorld") == "Hello\tWorld"
    )

    # Newline should be preserved
    test(
        "Newline preservation",
        sanitize_text("Hello\nWorld") == "Hello\nWorld"
    )

    # Carriage return should be preserved
    test(
        "Carriage return preservation",
        sanitize_text("Hello\rWorld") == "Hello\rWorld"
    )

    # Empty string handling
    test(
        "Empty string handling",
        sanitize_text("") == ""
    )

    # None handling
    test(
        "None handling",
        sanitize_text(None) is None
    )

    # Mixed control characters
    test(
        "Mixed control characters",
        sanitize_text("A\x00B\x01C\tD\nE\x0BF") == "ABC\tD\nEF"
    )


def test_xxe_protection():
    """Test that XXE attacks are blocked by defusedxml."""
    print("\nTesting XXE protection...")

    import defusedxml.ElementTree as ET
    from defusedxml import DefusedXmlException

    # Billion laughs attack (exponential entity expansion)
    billion_laughs = '''<?xml version="1.0"?>
<!DOCTYPE lolz [
  <!ENTITY lol "lol">
  <!ENTITY lol2 "&lol;&lol;&lol;&lol;&lol;&lol;&lol;&lol;&lol;&lol;">
  <!ENTITY lol3 "&lol2;&lol2;&lol2;&lol2;&lol2;&lol2;&lol2;&lol2;&lol2;&lol2;">
]>
<lolz>&lol3;</lolz>'''

    xxe_blocked = False
    try:
        ET.fromstring(billion_laughs)
    except DefusedXmlException:
        xxe_blocked = True
    except Exception as e:
        # Other exceptions also indicate the attack was blocked
        if "entities" in str(e).lower() or "dtd" in str(e).lower():
            xxe_blocked = True

    test("Billion laughs attack blocked", xxe_blocked)

    # External entity attack
    external_entity = '''<?xml version="1.0"?>
<!DOCTYPE foo [
  <!ENTITY xxe SYSTEM "file:///etc/passwd">
]>
<foo>&xxe;</foo>'''

    external_blocked = False
    try:
        ET.fromstring(external_entity)
    except DefusedXmlException:
        external_blocked = True
    except Exception as e:
        if "entities" in str(e).lower() or "dtd" in str(e).lower():
            external_blocked = True

    test("External entity attack blocked", external_blocked)


def test_malformed_docx_handling():
    """Test parser handles malformed DOCX files gracefully."""
    print("\nTesting malformed DOCX handling...")

    import subprocess
    import json

    parser_script = SCRIPT_DIR / "parse_docx.py"

    # Test with non-existent file
    result = subprocess.run(
        [sys.executable, str(parser_script), "/nonexistent/file.docx"],
        capture_output=True,
        text=True
    )
    test(
        "Non-existent file returns error",
        result.returncode != 0
    )

    # Test with empty file - parser returns JSON with warning
    with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as f:
        empty_file = f.name

    result = subprocess.run(
        [sys.executable, str(parser_script), empty_file],
        capture_output=True,
        text=True
    )
    empty_handled = False
    if result.returncode == 0:
        try:
            data = json.loads(result.stdout)
            warnings = data.get("stats", {}).get("warnings", [])
            empty_handled = any("Invalid ZIP" in w or "not a ZIP" in w.lower() for w in warnings)
        except json.JSONDecodeError:
            pass
    test(
        "Empty file produces warning",
        empty_handled,
        "Expected warning about invalid ZIP"
    )
    Path(empty_file).unlink()

    # Test with valid ZIP but not DOCX - parser returns JSON with warning
    with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as f:
        fake_docx = f.name

    with zipfile.ZipFile(fake_docx, 'w') as zf:
        zf.writestr("random.txt", "not a docx")

    result = subprocess.run(
        [sys.executable, str(parser_script), fake_docx],
        capture_output=True,
        text=True
    )
    fake_handled = False
    if result.returncode == 0:
        try:
            data = json.loads(result.stdout)
            warnings = data.get("stats", {}).get("warnings", [])
            # Should have warning about missing document.xml
            fake_handled = any("document.xml" in w.lower() or "not found" in w.lower() for w in warnings)
        except json.JSONDecodeError:
            pass
    test(
        "Fake DOCX (missing document.xml) produces warning",
        fake_handled,
        "Expected warning about missing document.xml"
    )
    Path(fake_docx).unlink()


def test_defusedxml_available():
    """Verify defusedxml is installed and working."""
    print("\nTesting defusedxml availability...")

    try:
        import defusedxml
        test("defusedxml installed", True)
        test(
            "defusedxml version available",
            hasattr(defusedxml, '__version__')
        )
    except ImportError:
        test("defusedxml installed", False, "pip3 install defusedxml")


def main():
    print("=" * 50)
    print("docx2pages Edge Case Tests")
    print("=" * 50)

    test_defusedxml_available()
    test_sanitize_text()
    test_xxe_protection()
    test_malformed_docx_handling()

    print()
    print("=" * 50)
    print(f"Results: {passed} passed, {failed} failed")
    print("=" * 50)

    sys.exit(0 if failed == 0 else 1)


if __name__ == "__main__":
    main()

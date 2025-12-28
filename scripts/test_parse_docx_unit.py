#!/usr/bin/env python3
"""
Unit tests for parse_docx.py functions.
Tests individual functions in isolation without requiring DOCX files.

Run with: python3 scripts/test_parse_docx_unit.py
"""

import unittest
import sys
from pathlib import Path

# Add scripts directory to path for imports
sys.path.insert(0, str(Path(__file__).parent))

from parse_docx import (
    sanitize_text,
    NAMESPACES,
    MAX_RECURSION_DEPTH,
)


class TestSanitizeText(unittest.TestCase):
    """Tests for sanitize_text() function."""

    def test_empty_string(self):
        """Empty string should return empty string."""
        self.assertEqual(sanitize_text(""), "")

    def test_none_input(self):
        """None input should return None."""
        self.assertIsNone(sanitize_text(None))

    def test_normal_text(self):
        """Normal text should pass through unchanged."""
        text = "Hello, world!"
        self.assertEqual(sanitize_text(text), text)

    def test_preserves_tab(self):
        """Tab character (0x09) should be preserved."""
        text = "Column1\tColumn2"
        self.assertEqual(sanitize_text(text), text)

    def test_preserves_newline(self):
        """Newline character (0x0A) should be preserved."""
        text = "Line1\nLine2"
        self.assertEqual(sanitize_text(text), text)

    def test_preserves_carriage_return(self):
        """Carriage return (0x0D) should be preserved."""
        text = "Line1\r\nLine2"
        self.assertEqual(sanitize_text(text), text)

    def test_removes_null_byte(self):
        """Null byte (0x00) should be removed."""
        text = "Hello\x00World"
        self.assertEqual(sanitize_text(text), "HelloWorld")

    def test_removes_control_characters(self):
        """Control characters (0x01-0x08, 0x0B, 0x0C, 0x0E-0x1F) should be removed."""
        # Test various control characters
        text = "A\x01B\x02C\x07D\x0BE\x0CF\x1FG"
        self.assertEqual(sanitize_text(text), "ABCDEFG")

    def test_mixed_content(self):
        """Mix of valid and invalid characters."""
        text = "Hello\x00\tWorld\x01\n!"
        self.assertEqual(sanitize_text(text), "Hello\tWorld\n!")

    def test_unicode_preserved(self):
        """Unicode characters should be preserved."""
        text = "Café résumé 日本語"
        self.assertEqual(sanitize_text(text), text)

    def test_emoji_preserved(self):
        """Emoji should be preserved."""
        text = "Hello 👋 World 🌍"
        self.assertEqual(sanitize_text(text), text)


class TestNamespaces(unittest.TestCase):
    """Tests for namespace constants."""

    def test_word_namespace_defined(self):
        """Word namespace should be defined."""
        self.assertIn('w', NAMESPACES)
        self.assertTrue(NAMESPACES['w'].startswith('http://'))

    def test_relationship_namespace_defined(self):
        """Relationship namespace should be defined."""
        self.assertIn('r', NAMESPACES)
        self.assertTrue(NAMESPACES['r'].startswith('http://'))


class TestRecursionLimit(unittest.TestCase):
    """Tests for recursion limit constant."""

    def test_recursion_limit_reasonable(self):
        """MAX_RECURSION_DEPTH should be a reasonable value."""
        self.assertGreaterEqual(MAX_RECURSION_DEPTH, 100)
        self.assertLessEqual(MAX_RECURSION_DEPTH, 10000)

    def test_system_recursion_limit_set(self):
        """System recursion limit should be at least MAX_RECURSION_DEPTH."""
        self.assertGreaterEqual(sys.getrecursionlimit(), MAX_RECURSION_DEPTH)


class TestGetText(unittest.TestCase):
    """Tests for get_text() function with mock XML elements."""

    def setUp(self):
        """Set up test fixtures."""
        import defusedxml.ElementTree as ET
        self.ET = ET

    def test_simple_text_element(self):
        """Extract text from simple w:t element."""
        from parse_docx import get_text
        xml = f'<w:t xmlns:w="{NAMESPACES["w"]}">Hello World</w:t>'
        elem = self.ET.fromstring(xml)
        self.assertEqual(get_text(elem), "Hello World")

    def test_nested_text(self):
        """Extract text from nested elements."""
        from parse_docx import get_text
        xml = f'''<w:r xmlns:w="{NAMESPACES["w"]}">
            <w:t>Hello </w:t>
            <w:t>World</w:t>
        </w:r>'''
        elem = self.ET.fromstring(xml)
        self.assertEqual(get_text(elem), "Hello World")

    def test_tab_element(self):
        """Tab elements should produce tab character."""
        from parse_docx import get_text
        xml = f'''<w:r xmlns:w="{NAMESPACES["w"]}">
            <w:t>A</w:t>
            <w:tab/>
            <w:t>B</w:t>
        </w:r>'''
        elem = self.ET.fromstring(xml)
        self.assertEqual(get_text(elem), "A\tB")

    def test_soft_break(self):
        """Soft line break (no type) should produce newline."""
        from parse_docx import get_text
        xml = f'''<w:r xmlns:w="{NAMESPACES["w"]}">
            <w:t>Line1</w:t>
            <w:br/>
            <w:t>Line2</w:t>
        </w:r>'''
        elem = self.ET.fromstring(xml)
        self.assertEqual(get_text(elem), "Line1\nLine2")

    def test_carriage_return(self):
        """Carriage return element should produce newline."""
        from parse_docx import get_text
        xml = f'''<w:r xmlns:w="{NAMESPACES["w"]}">
            <w:t>Line1</w:t>
            <w:cr/>
            <w:t>Line2</w:t>
        </w:r>'''
        elem = self.ET.fromstring(xml)
        self.assertEqual(get_text(elem), "Line1\nLine2")

    def test_page_break_ignored(self):
        """Page breaks (type=page) should be ignored."""
        from parse_docx import get_text
        xml = f'''<w:r xmlns:w="{NAMESPACES["w"]}">
            <w:t>Before</w:t>
            <w:br w:type="page"/>
            <w:t>After</w:t>
        </w:r>'''
        elem = self.ET.fromstring(xml)
        self.assertEqual(get_text(elem), "BeforeAfter")

    def test_extract_links_false(self):
        """With extract_links=False, should return just text."""
        from parse_docx import get_text
        xml = f'<w:t xmlns:w="{NAMESPACES["w"]}">Hello</w:t>'
        elem = self.ET.fromstring(xml)
        result = get_text(elem, extract_links=False)
        self.assertEqual(result, "Hello")
        self.assertIsInstance(result, str)

    def test_extract_links_true_no_links(self):
        """With extract_links=True but no links, should return (text, None)."""
        from parse_docx import get_text
        xml = f'<w:t xmlns:w="{NAMESPACES["w"]}">Hello</w:t>'
        elem = self.ET.fromstring(xml)
        text, links = get_text(elem, extract_links=True)
        self.assertEqual(text, "Hello")
        self.assertIsNone(links)


class TestParseTable(unittest.TestCase):
    """Tests for parse_table() function."""

    def setUp(self):
        """Set up test fixtures."""
        import defusedxml.ElementTree as ET
        self.ET = ET

    def test_simple_table(self):
        """Parse a simple 2x2 table."""
        from parse_docx import parse_table
        xml = f'''<w:tbl xmlns:w="{NAMESPACES["w"]}">
            <w:tr>
                <w:tc><w:p><w:r><w:t>A1</w:t></w:r></w:p></w:tc>
                <w:tc><w:p><w:r><w:t>B1</w:t></w:r></w:p></w:tc>
            </w:tr>
            <w:tr>
                <w:tc><w:p><w:r><w:t>A2</w:t></w:r></w:p></w:tc>
                <w:tc><w:p><w:r><w:t>B2</w:t></w:r></w:p></w:tc>
            </w:tr>
        </w:tbl>'''
        elem = self.ET.fromstring(xml)
        rows, merges = parse_table(elem)
        self.assertEqual(len(rows), 2)
        self.assertEqual(len(rows[0]), 2)
        self.assertEqual(rows[0][0], "A1")
        self.assertEqual(rows[0][1], "B1")
        self.assertEqual(rows[1][0], "A2")
        self.assertEqual(rows[1][1], "B2")
        self.assertIsNone(merges)

    def test_empty_table(self):
        """Parse a table with empty cells."""
        from parse_docx import parse_table
        xml = f'''<w:tbl xmlns:w="{NAMESPACES["w"]}">
            <w:tr>
                <w:tc><w:p/></w:tc>
                <w:tc><w:p><w:r><w:t>B1</w:t></w:r></w:p></w:tc>
            </w:tr>
        </w:tbl>'''
        elem = self.ET.fromstring(xml)
        rows, merges = parse_table(elem)
        self.assertEqual(len(rows), 1)
        self.assertEqual(rows[0][0], "")
        self.assertEqual(rows[0][1], "B1")


if __name__ == '__main__':
    # Run with verbose output
    unittest.main(verbosity=2)

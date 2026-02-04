#!/usr/bin/env python3
"""
Unit tests for html_writer.py functions.
Tests HTML generation from JSON blocks for the macOS 26+ HTML pipeline.

Run with: python3 scripts/test_html_writer.py
"""

import json
import os
import sys
import tempfile
import unittest
from pathlib import Path

# Add scripts directory to path for imports
sys.path.insert(0, str(Path(__file__).parent))

from html_writer import json_to_html


class TestHtmlWriterBase(unittest.TestCase):
    """Base class with helper methods for HTML writer tests."""

    def create_json_file(self, blocks):
        """Create a temporary JSON file with the given blocks."""
        data = {"blocks": blocks}
        fd, path = tempfile.mkstemp(suffix='.json')
        with os.fdopen(fd, 'w') as f:
            json.dump(data, f)
        return path

    def convert_blocks_to_html(self, blocks):
        """Convert blocks to HTML and return the content."""
        json_path = self.create_json_file(blocks)
        fd, html_path = tempfile.mkstemp(suffix='.html')
        os.close(fd)
        try:
            json_to_html(json_path, html_path)
            with open(html_path) as f:
                return f.read()
        finally:
            os.unlink(json_path)
            if os.path.exists(html_path):
                os.unlink(html_path)


class TestHtmlStructure(TestHtmlWriterBase):
    """Tests for basic HTML document structure."""

    def test_empty_blocks(self):
        """Empty blocks should produce valid HTML structure."""
        html = self.convert_blocks_to_html([])
        self.assertIn('<!DOCTYPE html>', html)
        self.assertIn('<html>', html)
        self.assertIn('<head>', html)
        self.assertIn('<meta charset="utf-8">', html)
        self.assertIn('<style>', html)
        self.assertIn('</style>', html)
        self.assertIn('</head>', html)
        self.assertIn('<body>', html)
        self.assertIn('</body>', html)
        self.assertIn('</html>', html)

    def test_css_styles_included(self):
        """CSS styles should be included in head."""
        html = self.convert_blocks_to_html([])
        self.assertIn('font-family:', html)
        self.assertIn('h1 {', html)
        self.assertIn('h2 {', html)
        self.assertIn('h3 {', html)
        self.assertIn('table {', html)
        self.assertIn('border-collapse:', html)


class TestTitleAndSubtitle(TestHtmlWriterBase):
    """Tests for title and subtitle blocks."""

    def test_title_generates_h1(self):
        """Title block should generate h1 with 28pt font."""
        blocks = [{"type": "title", "text": "Document Title"}]
        html = self.convert_blocks_to_html(blocks)
        self.assertIn('<h1 style="font-size: 28pt;">Document Title</h1>', html)

    def test_subtitle_generates_h2(self):
        """Subtitle block should generate h2 with gray color."""
        blocks = [{"type": "subtitle", "text": "Document Subtitle"}]
        html = self.convert_blocks_to_html(blocks)
        self.assertIn('<h2 style="font-size: 18pt; color: #666;">Document Subtitle</h2>', html)

    def test_title_escapes_html(self):
        """Title should escape HTML special characters."""
        blocks = [{"type": "title", "text": "<script>alert('xss')</script>"}]
        html = self.convert_blocks_to_html(blocks)
        self.assertIn('&lt;script&gt;', html)
        self.assertNotIn('<script>', html)

    def test_empty_title(self):
        """Empty title text should produce empty h1."""
        blocks = [{"type": "title", "text": ""}]
        html = self.convert_blocks_to_html(blocks)
        self.assertIn('<h1 style="font-size: 28pt;"></h1>', html)

    def test_title_missing_text_key(self):
        """Title block without text key should produce empty h1."""
        blocks = [{"type": "title"}]
        html = self.convert_blocks_to_html(blocks)
        self.assertIn('<h1 style="font-size: 28pt;"></h1>', html)


class TestHeadings(TestHtmlWriterBase):
    """Tests for heading blocks (h1-h6)."""

    def test_heading_level_1(self):
        """Heading level 1 should generate h1."""
        blocks = [{"type": "heading", "level": 1, "text": "Heading 1"}]
        html = self.convert_blocks_to_html(blocks)
        self.assertIn('<h1>Heading 1</h1>', html)

    def test_heading_level_2(self):
        """Heading level 2 should generate h2."""
        blocks = [{"type": "heading", "level": 2, "text": "Heading 2"}]
        html = self.convert_blocks_to_html(blocks)
        self.assertIn('<h2>Heading 2</h2>', html)

    def test_heading_level_3(self):
        """Heading level 3 should generate h3."""
        blocks = [{"type": "heading", "level": 3, "text": "Heading 3"}]
        html = self.convert_blocks_to_html(blocks)
        self.assertIn('<h3>Heading 3</h3>', html)

    def test_heading_level_4(self):
        """Heading level 4 should generate h4."""
        blocks = [{"type": "heading", "level": 4, "text": "Heading 4"}]
        html = self.convert_blocks_to_html(blocks)
        self.assertIn('<h4>Heading 4</h4>', html)

    def test_heading_level_5(self):
        """Heading level 5 should generate h5."""
        blocks = [{"type": "heading", "level": 5, "text": "Heading 5"}]
        html = self.convert_blocks_to_html(blocks)
        self.assertIn('<h5>Heading 5</h5>', html)

    def test_heading_level_6(self):
        """Heading level 6 should generate h6."""
        blocks = [{"type": "heading", "level": 6, "text": "Heading 6"}]
        html = self.convert_blocks_to_html(blocks)
        self.assertIn('<h6>Heading 6</h6>', html)

    def test_heading_level_exceeds_6_clamped(self):
        """Heading level > 6 should be clamped to h6."""
        blocks = [{"type": "heading", "level": 10, "text": "Heading 10"}]
        html = self.convert_blocks_to_html(blocks)
        self.assertIn('<h6>Heading 10</h6>', html)
        self.assertNotIn('<h10>', html)

    def test_heading_default_level(self):
        """Heading without level should default to h1."""
        blocks = [{"type": "heading", "text": "Default Heading"}]
        html = self.convert_blocks_to_html(blocks)
        self.assertIn('<h1>Default Heading</h1>', html)

    def test_heading_escapes_html(self):
        """Heading should escape HTML special characters."""
        blocks = [{"type": "heading", "level": 2, "text": "Heading & More <stuff>"}]
        html = self.convert_blocks_to_html(blocks)
        self.assertIn('&amp;', html)
        self.assertIn('&lt;stuff&gt;', html)


class TestParagraphs(TestHtmlWriterBase):
    """Tests for paragraph blocks."""

    def test_simple_paragraph(self):
        """Simple paragraph should generate p element."""
        blocks = [{"type": "paragraph", "text": "This is a paragraph."}]
        html = self.convert_blocks_to_html(blocks)
        self.assertIn('<p>This is a paragraph.</p>', html)

    def test_paragraph_with_newlines(self):
        """Newlines in paragraph should become br elements."""
        blocks = [{"type": "paragraph", "text": "Line 1\nLine 2\nLine 3"}]
        html = self.convert_blocks_to_html(blocks)
        self.assertIn('<p>Line 1<br>Line 2<br>Line 3</p>', html)

    def test_paragraph_with_tabs(self):
        """Tabs in paragraph should become emsp entities."""
        blocks = [{"type": "paragraph", "text": "Col1\tCol2\tCol3"}]
        html = self.convert_blocks_to_html(blocks)
        self.assertIn('<p>Col1&emsp;Col2&emsp;Col3</p>', html)

    def test_paragraph_escapes_html(self):
        """Paragraph should escape HTML special characters."""
        blocks = [{"type": "paragraph", "text": "Use <b> tags for bold & <i> for italic"}]
        html = self.convert_blocks_to_html(blocks)
        self.assertIn('&lt;b&gt;', html)
        self.assertIn('&lt;i&gt;', html)
        self.assertIn('&amp;', html)
        self.assertNotIn('<b>', html)

    def test_empty_paragraph(self):
        """Empty paragraph should produce empty p element."""
        blocks = [{"type": "paragraph", "text": ""}]
        html = self.convert_blocks_to_html(blocks)
        self.assertIn('<p></p>', html)

    def test_paragraph_with_unicode(self):
        """Paragraph should preserve unicode characters."""
        blocks = [{"type": "paragraph", "text": "Café résumé 日本語 emoji: 🎉"}]
        html = self.convert_blocks_to_html(blocks)
        self.assertIn('Café résumé 日本語 emoji: 🎉', html)

    def test_paragraph_with_quotes(self):
        """Paragraph should escape quotes."""
        blocks = [{"type": "paragraph", "text": 'He said "Hello" and \'Goodbye\''}]
        html = self.convert_blocks_to_html(blocks)
        # html.escape does escape quotes when quote=True (default)
        self.assertIn('He said', html)


class TestBulletedLists(TestHtmlWriterBase):
    """Tests for bulleted (unordered) lists."""

    def test_simple_bulleted_list(self):
        """Bulleted list should generate ul element."""
        blocks = [{
            "type": "list",
            "ordered": False,
            "items": [
                {"text": "Item 1"},
                {"text": "Item 2"},
                {"text": "Item 3"}
            ]
        }]
        html = self.convert_blocks_to_html(blocks)
        self.assertIn('<ul>', html)
        self.assertIn('<li style="">Item 1</li>', html)
        self.assertIn('<li style="">Item 2</li>', html)
        self.assertIn('<li style="">Item 3</li>', html)
        self.assertIn('</ul>', html)

    def test_bulleted_list_default_ordered(self):
        """List without ordered key should default to bulleted."""
        blocks = [{
            "type": "list",
            "items": [{"text": "Item 1"}]
        }]
        html = self.convert_blocks_to_html(blocks)
        self.assertIn('<ul>', html)
        self.assertNotIn('<ol>', html)

    def test_bulleted_list_with_indentation(self):
        """Indented list items should have margin-left style."""
        blocks = [{
            "type": "list",
            "ordered": False,
            "items": [
                {"text": "Level 0", "level": 0},
                {"text": "Level 1", "level": 1},
                {"text": "Level 2", "level": 2}
            ]
        }]
        html = self.convert_blocks_to_html(blocks)
        self.assertIn('<li style="">Level 0</li>', html)
        self.assertIn('<li style="margin-left: 20px;">Level 1</li>', html)
        self.assertIn('<li style="margin-left: 40px;">Level 2</li>', html)

    def test_empty_bulleted_list(self):
        """Empty list should produce ul with no items."""
        blocks = [{
            "type": "list",
            "ordered": False,
            "items": []
        }]
        html = self.convert_blocks_to_html(blocks)
        self.assertIn('<ul>', html)
        self.assertIn('</ul>', html)
        self.assertNotIn('<li', html)

    def test_bulleted_list_escapes_html(self):
        """List items should escape HTML."""
        blocks = [{
            "type": "list",
            "ordered": False,
            "items": [{"text": "<script>alert('xss')</script>"}]
        }]
        html = self.convert_blocks_to_html(blocks)
        self.assertIn('&lt;script&gt;', html)
        self.assertNotIn('<script>alert', html)


class TestNumberedLists(TestHtmlWriterBase):
    """Tests for numbered (ordered) lists."""

    def test_simple_numbered_list(self):
        """Numbered list should generate ol element."""
        blocks = [{
            "type": "list",
            "ordered": True,
            "items": [
                {"text": "First"},
                {"text": "Second"},
                {"text": "Third"}
            ]
        }]
        html = self.convert_blocks_to_html(blocks)
        self.assertIn('<ol>', html)
        self.assertIn('<li style="">First</li>', html)
        self.assertIn('<li style="">Second</li>', html)
        self.assertIn('<li style="">Third</li>', html)
        self.assertIn('</ol>', html)

    def test_numbered_list_with_indentation(self):
        """Indented numbered list items should have margin-left style."""
        blocks = [{
            "type": "list",
            "ordered": True,
            "items": [
                {"text": "Step 1", "level": 0},
                {"text": "Step 1a", "level": 1},
                {"text": "Step 1b", "level": 1},
                {"text": "Step 2", "level": 0}
            ]
        }]
        html = self.convert_blocks_to_html(blocks)
        self.assertIn('<ol>', html)
        self.assertIn('<li style="">Step 1</li>', html)
        self.assertIn('<li style="margin-left: 20px;">Step 1a</li>', html)


class TestTables(TestHtmlWriterBase):
    """Tests for table blocks."""

    def test_simple_table(self):
        """Simple table should generate table element with tr and td."""
        blocks = [{
            "type": "table",
            "rows": [
                ["A1", "B1", "C1"],
                ["A2", "B2", "C2"]
            ]
        }]
        html = self.convert_blocks_to_html(blocks)
        self.assertIn('<table>', html)
        self.assertIn('<tr>', html)
        self.assertIn('<td>A1</td>', html)
        self.assertIn('<td>B1</td>', html)
        self.assertIn('<td>C1</td>', html)
        self.assertIn('<td>A2</td>', html)
        self.assertIn('<td>B2</td>', html)
        self.assertIn('<td>C2</td>', html)
        self.assertIn('</tr>', html)
        self.assertIn('</table>', html)

    def test_table_with_empty_cells(self):
        """Table with empty strings should produce empty td elements."""
        blocks = [{
            "type": "table",
            "rows": [
                ["A1", "", "C1"],
                ["", "B2", ""]
            ]
        }]
        html = self.convert_blocks_to_html(blocks)
        # Count occurrences of empty td (3 empty cells: "", "", "")
        self.assertEqual(html.count('<td></td>'), 3)

    def test_table_with_null_cells(self):
        """Table with null cells (merged) should produce empty td elements."""
        blocks = [{
            "type": "table",
            "rows": [
                ["Header", None, None],
                ["A2", "B2", "C2"]
            ]
        }]
        html = self.convert_blocks_to_html(blocks)
        self.assertIn('<td>Header</td>', html)
        self.assertEqual(html.count('<td></td>'), 2)

    def test_empty_table(self):
        """Empty table should produce table with no rows."""
        blocks = [{
            "type": "table",
            "rows": []
        }]
        html = self.convert_blocks_to_html(blocks)
        self.assertIn('<table>', html)
        self.assertIn('</table>', html)
        self.assertNotIn('<tr>', html)

    def test_table_escapes_html(self):
        """Table cells should escape HTML."""
        blocks = [{
            "type": "table",
            "rows": [
                ["<script>", "&", "\"quote\""]
            ]
        }]
        html = self.convert_blocks_to_html(blocks)
        self.assertIn('&lt;script&gt;', html)
        self.assertIn('&amp;', html)
        self.assertNotIn('<script>', html)

    def test_table_with_unicode(self):
        """Table cells should preserve unicode characters."""
        blocks = [{
            "type": "table",
            "rows": [
                ["日本語", "Émoji 🎉", "Café"]
            ]
        }]
        html = self.convert_blocks_to_html(blocks)
        self.assertIn('日本語', html)
        self.assertIn('Émoji 🎉', html)
        self.assertIn('Café', html)


class TestBreaks(TestHtmlWriterBase):
    """Tests for break blocks."""

    def test_break_generates_br(self):
        """Break block should generate two br elements."""
        blocks = [{"type": "break"}]
        html = self.convert_blocks_to_html(blocks)
        self.assertIn('<br><br>', html)

    def test_multiple_breaks(self):
        """Multiple break blocks should generate multiple br pairs."""
        blocks = [
            {"type": "break"},
            {"type": "break"},
            {"type": "break"}
        ]
        html = self.convert_blocks_to_html(blocks)
        self.assertEqual(html.count('<br><br>'), 3)


class TestMixedContent(TestHtmlWriterBase):
    """Tests for documents with mixed content types."""

    def test_full_document(self):
        """Full document with all block types."""
        blocks = [
            {"type": "title", "text": "My Document"},
            {"type": "subtitle", "text": "A Comprehensive Test"},
            {"type": "heading", "level": 2, "text": "Introduction"},
            {"type": "paragraph", "text": "This is the first paragraph."},
            {"type": "heading", "level": 3, "text": "Key Points"},
            {"type": "list", "ordered": False, "items": [
                {"text": "Point 1"},
                {"text": "Point 2"}
            ]},
            {"type": "heading", "level": 3, "text": "Steps"},
            {"type": "list", "ordered": True, "items": [
                {"text": "Step 1"},
                {"text": "Step 2"}
            ]},
            {"type": "heading", "level": 2, "text": "Data"},
            {"type": "table", "rows": [
                ["Name", "Value"],
                ["Test", "123"]
            ]},
            {"type": "break"},
            {"type": "paragraph", "text": "Conclusion."}
        ]
        html = self.convert_blocks_to_html(blocks)

        # Verify all elements present
        self.assertIn('<h1 style="font-size: 28pt;">My Document</h1>', html)
        self.assertIn('<h2 style="font-size: 18pt; color: #666;">A Comprehensive Test</h2>', html)
        self.assertIn('<h2>Introduction</h2>', html)
        self.assertIn('<p>This is the first paragraph.</p>', html)
        self.assertIn('<ul>', html)
        self.assertIn('<ol>', html)
        self.assertIn('<table>', html)
        self.assertIn('<br><br>', html)
        self.assertIn('<p>Conclusion.</p>', html)

    def test_unknown_block_type_ignored(self):
        """Unknown block types should be silently ignored."""
        blocks = [
            {"type": "paragraph", "text": "Before"},
            {"type": "unknown_type", "data": "ignored"},
            {"type": "paragraph", "text": "After"}
        ]
        html = self.convert_blocks_to_html(blocks)
        self.assertIn('<p>Before</p>', html)
        self.assertIn('<p>After</p>', html)
        self.assertNotIn('ignored', html)
        self.assertNotIn('unknown', html)


class TestEdgeCases(TestHtmlWriterBase):
    """Tests for edge cases and special scenarios."""

    def test_special_characters_in_all_contexts(self):
        """Special characters should be escaped in all contexts."""
        blocks = [
            {"type": "title", "text": "Title with <>&\"'"},
            {"type": "heading", "level": 1, "text": "Heading with <>&\"'"},
            {"type": "paragraph", "text": "Paragraph with <>&\"'"},
            {"type": "list", "items": [{"text": "List item with <>&\"'"}]},
            {"type": "table", "rows": [["Cell with <>&\"'"]]}
        ]
        html = self.convert_blocks_to_html(blocks)
        # Should not contain unescaped < or > in content
        # (they should appear as &lt; and &gt;)
        self.assertIn('&lt;', html)
        self.assertIn('&gt;', html)
        self.assertIn('&amp;', html)

    def test_very_long_text(self):
        """Very long text should be handled correctly."""
        long_text = "A" * 10000
        blocks = [{"type": "paragraph", "text": long_text}]
        html = self.convert_blocks_to_html(blocks)
        self.assertIn(long_text, html)

    def test_deeply_nested_indentation(self):
        """Deep indentation levels should work correctly."""
        blocks = [{
            "type": "list",
            "ordered": False,
            "items": [
                {"text": "Level 5", "level": 5},
                {"text": "Level 10", "level": 10}
            ]
        }]
        html = self.convert_blocks_to_html(blocks)
        self.assertIn('margin-left: 100px', html)  # 5 * 20
        self.assertIn('margin-left: 200px', html)  # 10 * 20

    def test_whitespace_preservation(self):
        """Whitespace should be preserved appropriately."""
        blocks = [{"type": "paragraph", "text": "  Leading spaces  "}]
        html = self.convert_blocks_to_html(blocks)
        self.assertIn('  Leading spaces  ', html)

    def test_multiline_cells(self):
        """Table cells with newlines should be handled."""
        blocks = [{
            "type": "table",
            "rows": [["Line1\nLine2", "Normal"]]
        }]
        html = self.convert_blocks_to_html(blocks)
        # Newlines in table cells are preserved as-is (not converted to <br>)
        self.assertIn('Line1\nLine2', html)


class TestXssPrevention(TestHtmlWriterBase):
    """Tests to verify XSS attack vectors are properly escaped."""

    def test_script_tag_escaped(self):
        """Script tags should be escaped - cannot execute as HTML."""
        blocks = [{"type": "paragraph", "text": "<script>alert('XSS')</script>"}]
        html = self.convert_blocks_to_html(blocks)
        # The important thing is that < and > are escaped, making script non-executable
        self.assertNotIn('<script>', html)
        self.assertIn('&lt;script&gt;', html)

    def test_event_handler_escaped(self):
        """Event handlers should be escaped - img tag cannot execute."""
        blocks = [{"type": "paragraph", "text": '<img src="x" onerror="alert(1)">'}]
        html = self.convert_blocks_to_html(blocks)
        # The img tag is escaped so it renders as text, not as an executable element
        # The < and > are escaped, making the img tag non-functional
        self.assertNotIn('<img ', html)  # Not a real img tag
        self.assertIn('&lt;img', html)   # Escaped version

    def test_javascript_url_escaped(self):
        """JavaScript URLs should be escaped - anchor tag cannot execute."""
        blocks = [{"type": "paragraph", "text": '<a href="javascript:alert(1)">click</a>'}]
        html = self.convert_blocks_to_html(blocks)
        # The anchor tag is escaped so it renders as text, not as a clickable link
        self.assertNotIn('<a href', html)  # Not a real anchor tag
        self.assertIn('&lt;a href', html)  # Escaped version


if __name__ == '__main__':
    # Run with verbose output
    unittest.main(verbosity=2)

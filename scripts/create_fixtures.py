#!/usr/bin/env python3
"""
Create test DOCX fixtures for docx2pages testing.
"""

import os
import zipfile
from pathlib import Path

FIXTURES_DIR = Path(__file__).parent.parent / "fixtures"

# Word XML templates
CONTENT_TYPES_XML = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
  <Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>
</Types>'''

RELS_XML = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>'''

DOCUMENT_RELS_XML = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>
</Relationships>'''

STYLES_XML = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:style w:type="paragraph" w:styleId="Normal" w:default="1">
    <w:name w:val="Normal"/>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Title">
    <w:name w:val="Title"/>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Subtitle">
    <w:name w:val="Subtitle"/>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Heading1">
    <w:name w:val="Heading 1"/>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Heading2">
    <w:name w:val="Heading 2"/>
    <w:basedOn w:val="Heading1"/>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Heading3">
    <w:name w:val="Heading 3"/>
    <w:basedOn w:val="Heading2"/>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Heading4">
    <w:name w:val="Heading 4"/>
    <w:basedOn w:val="Heading3"/>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Heading5">
    <w:name w:val="Heading 5"/>
    <w:basedOn w:val="Heading4"/>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Heading6">
    <w:name w:val="Heading 6"/>
    <w:basedOn w:val="Heading5"/>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Heading7">
    <w:name w:val="Heading 7"/>
    <w:basedOn w:val="Heading6"/>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Heading8">
    <w:name w:val="Heading 8"/>
    <w:basedOn w:val="Heading7"/>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Heading9">
    <w:name w:val="Heading 9"/>
    <w:basedOn w:val="Heading8"/>
  </w:style>
  <w:style w:type="paragraph" w:styleId="ListParagraph">
    <w:name w:val="List Paragraph"/>
  </w:style>
  <w:style w:type="paragraph" w:styleId="CustomHeading">
    <w:name w:val="Custom Heading"/>
    <w:basedOn w:val="Heading2"/>
  </w:style>
</w:styles>'''

NUMBERING_XML = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:abstractNum w:abstractNumId="0">
    <w:lvl w:ilvl="0">
      <w:numFmt w:val="bullet"/>
    </w:lvl>
    <w:lvl w:ilvl="1">
      <w:numFmt w:val="bullet"/>
    </w:lvl>
    <w:lvl w:ilvl="2">
      <w:numFmt w:val="bullet"/>
    </w:lvl>
  </w:abstractNum>
  <w:abstractNum w:abstractNumId="1">
    <w:lvl w:ilvl="0">
      <w:numFmt w:val="decimal"/>
    </w:lvl>
    <w:lvl w:ilvl="1">
      <w:numFmt w:val="lowerLetter"/>
    </w:lvl>
    <w:lvl w:ilvl="2">
      <w:numFmt w:val="lowerRoman"/>
    </w:lvl>
  </w:abstractNum>
  <w:num w:numId="1">
    <w:abstractNumId w:val="0"/>
  </w:num>
  <w:num w:numId="2">
    <w:abstractNumId w:val="1"/>
  </w:num>
</w:numbering>'''

def escape_xml(text):
    """Escape special XML characters."""
    return (text
            .replace('&', '&amp;')
            .replace('<', '&lt;')
            .replace('>', '&gt;')
            .replace('"', '&quot;')
            .replace("'", '&apos;'))

def make_paragraph(text, style=None):
    """Create a paragraph XML element."""
    style_xml = f'<w:pPr><w:pStyle w:val="{style}"/></w:pPr>' if style else ''
    return f'<w:p>{style_xml}<w:r><w:t>{escape_xml(text)}</w:t></w:r></w:p>'

def make_paragraph_with_whitespace(runs):
    """
    Create a paragraph with tabs and line breaks.
    runs is a list of dicts: {'type': 'text'|'tab'|'br', 'value': '...' for text}
    """
    run_xml = []
    for run in runs:
        if run['type'] == 'text':
            run_xml.append(f'<w:r><w:t>{escape_xml(run["value"])}</w:t></w:r>')
        elif run['type'] == 'tab':
            run_xml.append('<w:r><w:tab/></w:r>')
        elif run['type'] == 'br':
            run_xml.append('<w:r><w:br/></w:r>')
        elif run['type'] == 'cr':
            run_xml.append('<w:r><w:cr/></w:r>')
    return f'<w:p>{"".join(run_xml)}</w:p>'

def make_list_item(text, num_id, ilvl=0):
    """Create a list item paragraph."""
    return f'''<w:p>
      <w:pPr>
        <w:pStyle w:val="ListParagraph"/>
        <w:numPr>
          <w:ilvl w:val="{ilvl}"/>
          <w:numId w:val="{num_id}"/>
        </w:numPr>
      </w:pPr>
      <w:r><w:t>{escape_xml(text)}</w:t></w:r>
    </w:p>'''

def make_table(rows):
    """Create a table XML element."""
    table_rows = []
    for row in rows:
        cells = ''.join(f'<w:tc><w:p><w:r><w:t>{escape_xml(str(cell))}</w:t></w:r></w:p></w:tc>' for cell in row)
        table_rows.append(f'<w:tr>{cells}</w:tr>')
    return f'<w:tbl><w:tblPr/><w:tblGrid/>{" ".join(table_rows)}</w:tbl>'

def make_document(body_content):
    """Create the document.xml content."""
    return f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    {body_content}
  </w:body>
</w:document>'''

def create_docx(filename, body_content):
    """Create a DOCX file with the given body content."""
    filepath = FIXTURES_DIR / filename

    with zipfile.ZipFile(filepath, 'w', zipfile.ZIP_DEFLATED) as zf:
        zf.writestr('[Content_Types].xml', CONTENT_TYPES_XML)
        zf.writestr('_rels/.rels', RELS_XML)
        zf.writestr('word/_rels/document.xml.rels', DOCUMENT_RELS_XML)
        zf.writestr('word/styles.xml', STYLES_XML)
        zf.writestr('word/numbering.xml', NUMBERING_XML)
        zf.writestr('word/document.xml', make_document(body_content))

    print(f"Created: {filepath}")

def create_all_headings_fixture():
    """Fixture 1: All heading levels 1-9 plus Title and Subtitle."""
    body = '\n'.join([
        make_paragraph("Document Title", "Title"),
        make_paragraph("Document Subtitle", "Subtitle"),
        make_paragraph("This is an introductory paragraph.", "Normal"),
        make_paragraph("Chapter One", "Heading1"),
        make_paragraph("Some content under chapter one.", "Normal"),
        make_paragraph("Section 1.1", "Heading2"),
        make_paragraph("Content under section 1.1.", "Normal"),
        make_paragraph("Subsection 1.1.1", "Heading3"),
        make_paragraph("Content under subsection.", "Normal"),
        make_paragraph("Level 4 Heading", "Heading4"),
        make_paragraph("Content under level 4.", "Normal"),
        make_paragraph("Level 5 Heading", "Heading5"),
        make_paragraph("Content under level 5.", "Normal"),
        make_paragraph("Level 6 Heading", "Heading6"),
        make_paragraph("Content under level 6.", "Normal"),
        make_paragraph("Level 7 Heading", "Heading7"),
        make_paragraph("Content under level 7.", "Normal"),
        make_paragraph("Level 8 Heading", "Heading8"),
        make_paragraph("Content under level 8.", "Normal"),
        make_paragraph("Level 9 Heading", "Heading9"),
        make_paragraph("Content under level 9.", "Normal"),
        make_paragraph("Custom Derived Heading", "CustomHeading"),
        make_paragraph("This uses a custom style derived from Heading 2.", "Normal"),
    ])
    create_docx("all_headings.docx", body)

def create_mixed_lists_fixture():
    """Fixture 2: Mixed bulleted and numbered lists with nesting."""
    body = '\n'.join([
        make_paragraph("Document with Mixed Lists", "Heading1"),
        make_paragraph("Here is a bulleted list:", "Normal"),
        # Bulleted list (numId=1)
        make_list_item("First bullet item", 1, 0),
        make_list_item("Second bullet item", 1, 0),
        make_list_item("Nested bullet item", 1, 1),
        make_list_item("Another nested item", 1, 1),
        make_list_item("Deeply nested item", 1, 2),
        make_list_item("Third bullet item", 1, 0),
        make_paragraph("And here is a numbered list:", "Normal"),
        # Numbered list (numId=2)
        make_list_item("First numbered item", 2, 0),
        make_list_item("Second numbered item", 2, 0),
        make_list_item("Sub-item a", 2, 1),
        make_list_item("Sub-item b", 2, 1),
        make_list_item("Sub-sub-item i", 2, 2),
        make_list_item("Third numbered item", 2, 0),
        make_paragraph("End of lists section.", "Normal"),
    ])
    create_docx("mixed_lists.docx", body)

def create_tables_fixture():
    """Fixture 3: Multiple tables with different sizes."""
    body = '\n'.join([
        make_paragraph("Document with Tables", "Heading1"),
        make_paragraph("Simple 3x3 Table", "Heading2"),
        make_paragraph("Here is a simple table:", "Normal"),
        make_table([
            ["Header 1", "Header 2", "Header 3"],
            ["Row 1 Col 1", "Row 1 Col 2", "Row 1 Col 3"],
            ["Row 2 Col 1", "Row 2 Col 2", "Row 2 Col 3"],
        ]),
        make_paragraph("Data Table", "Heading2"),
        make_paragraph("A larger data table:", "Normal"),
        make_table([
            ["Name", "Age", "City", "Country"],
            ["Alice", "30", "New York", "USA"],
            ["Bob", "25", "London", "UK"],
            ["Charlie", "35", "Paris", "France"],
            ["Diana", "28", "Tokyo", "Japan"],
        ]),
        make_paragraph("Conclusion", "Heading2"),
        make_paragraph("Tables demonstrate structure preservation.", "Normal"),
    ])
    create_docx("tables.docx", body)

def create_comprehensive_fixture():
    """Fixture 4: Comprehensive document with all element types."""
    body = '\n'.join([
        make_paragraph("Comprehensive Test Document", "Title"),
        make_paragraph("Testing all supported elements", "Subtitle"),

        make_paragraph("Introduction", "Heading1"),
        make_paragraph("This document tests all supported document elements including headings, paragraphs, lists, and tables.", "Normal"),

        make_paragraph("Document Structure", "Heading2"),
        make_paragraph("The following sections demonstrate each element type.", "Normal"),

        make_paragraph("Heading Hierarchy", "Heading3"),
        make_paragraph("We support all standard heading levels.", "Normal"),

        make_paragraph("Lists Section", "Heading2"),
        make_paragraph("Bulleted items:", "Normal"),
        make_list_item("Feature one", 1, 0),
        make_list_item("Feature two", 1, 0),
        make_list_item("Sub-feature", 1, 1),

        make_paragraph("Numbered steps:", "Normal"),
        make_list_item("First step", 2, 0),
        make_list_item("Second step", 2, 0),
        make_list_item("Third step", 2, 0),

        make_paragraph("Tables Section", "Heading2"),
        make_table([
            ["Element", "Supported", "Notes"],
            ["Headings", "Yes", "Levels 1-9"],
            ["Lists", "Yes", "Bulleted and numbered"],
            ["Tables", "Yes", "Basic structure"],
            ["Inline styles", "No", "Dropped by design"],
        ]),

        make_paragraph("Deep Headings", "Heading2"),
        make_paragraph("Testing deeper heading levels:", "Normal"),
        make_paragraph("Level 4 Section", "Heading4"),
        make_paragraph("Content at level 4.", "Normal"),
        make_paragraph("Level 5 Section", "Heading5"),
        make_paragraph("Content at level 5.", "Normal"),
        make_paragraph("Level 6 Section", "Heading6"),
        make_paragraph("Content at level 6.", "Normal"),

        make_paragraph("Conclusion", "Heading1"),
        make_paragraph("This concludes the comprehensive test document.", "Normal"),
    ])
    create_docx("comprehensive.docx", body)

def create_large_fixture():
    """Fixture 5: Large document for performance testing (300+ paragraphs)."""
    body_parts = []

    body_parts.append(make_paragraph("Large Performance Test Document", "Title"))
    body_parts.append(make_paragraph("Testing scalability with many elements", "Subtitle"))

    paragraph_count = 0

    # Create 10 chapters with sections
    for chapter in range(1, 11):
        body_parts.append(make_paragraph(f"Chapter {chapter}: Lorem Ipsum Section", "Heading1"))
        paragraph_count += 1

        for section in range(1, 6):
            body_parts.append(make_paragraph(f"Section {chapter}.{section}", "Heading2"))
            paragraph_count += 1

            # Add 5 paragraphs per section
            for para in range(5):
                body_parts.append(make_paragraph(
                    f"This is paragraph {para + 1} in section {chapter}.{section}. "
                    "Lorem ipsum dolor sit amet, consectetur adipiscing elit. "
                    "Sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.",
                    "Normal"
                ))
                paragraph_count += 1

            # Add a list every other section
            if section % 2 == 0:
                body_parts.append(make_paragraph("Key points:", "Normal"))
                for item in range(3):
                    body_parts.append(make_list_item(f"Point {item + 1} for section {chapter}.{section}", 1, 0))

            # Add subsection
            body_parts.append(make_paragraph(f"Subsection {chapter}.{section}.1", "Heading3"))
            paragraph_count += 1

            for para in range(3):
                body_parts.append(make_paragraph(
                    f"Subsection content paragraph {para + 1}. "
                    "Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris.",
                    "Normal"
                ))
                paragraph_count += 1

    # Add 3 tables
    body_parts.append(make_paragraph("Data Tables", "Heading1"))

    for table_num in range(1, 4):
        body_parts.append(make_paragraph(f"Table {table_num}: Sample Data", "Heading2"))
        rows = [["Column A", "Column B", "Column C", "Column D", "Column E"]]
        for row in range(5):
            rows.append([f"R{row+1}C{c+1}" for c in range(5)])
        body_parts.append(make_table(rows))
        body_parts.append(make_paragraph(f"Analysis of table {table_num} data follows.", "Normal"))

    # Add more numbered lists
    body_parts.append(make_paragraph("Summary Steps", "Heading1"))
    for i in range(10):
        body_parts.append(make_list_item(f"Summary step {i + 1}: Important action item", 2, 0))
        if i % 3 == 0:
            body_parts.append(make_list_item("Sub-step detail", 2, 1))

    body_parts.append(make_paragraph("Conclusion", "Heading1"))
    body_parts.append(make_paragraph(
        "This large document tests performance with 300+ paragraphs, "
        "multiple heading levels, nested lists, and multiple tables.",
        "Normal"
    ))

    body = '\n'.join(body_parts)
    create_docx("large.docx", body)
    print(f"  (Contains approximately {paragraph_count}+ paragraphs)")

def create_wide_table_fixture():
    """Fixture 6: Table with 30+ columns to test column addressing beyond Z."""
    body_parts = []

    body_parts.append(make_paragraph("Wide Table Test", "Title"))
    body_parts.append(make_paragraph("Testing table column addressing beyond 26 columns", "Subtitle"))

    body_parts.append(make_paragraph("Table with 35 Columns", "Heading1"))
    body_parts.append(make_paragraph(
        "This table has 35 columns (A through AI) to test Excel-style column addressing.",
        "Normal"
    ))

    # Create header row with column names A-AI (35 columns)
    num_cols = 35
    header_row = []
    for i in range(num_cols):
        # Generate Excel-style column names: A, B, ..., Z, AA, AB, ..., AI
        if i < 26:
            header_row.append(chr(65 + i))
        else:
            header_row.append('A' + chr(65 + (i - 26)))

    # Create 3 data rows
    rows = [header_row]
    for row_num in range(1, 4):
        row = [f"R{row_num}C{i+1}" for i in range(num_cols)]
        rows.append(row)

    body_parts.append(make_table(rows))

    body_parts.append(make_paragraph("Verification", "Heading2"))
    body_parts.append(make_paragraph(
        "The table above should have columns A through AI (35 total). "
        "Columns AA through AI verify addressing beyond the first 26 letters.",
        "Normal"
    ))

    body = '\n'.join(body_parts)
    create_docx("wide_table.docx", body)

def create_empty_fixture():
    """Fixture 7: Empty document with no content (edge case)."""
    # Empty body - tests handling of documents with no paragraphs
    create_docx("empty.docx", "")

def create_minimal_fixture():
    """Fixture 8: Minimal document with just one paragraph (edge case)."""
    body = make_paragraph("This is the only paragraph in this document.", "Normal")
    create_docx("minimal.docx", body)

def create_whitespace_fixture():
    """Fixture 9: Document with tabs and soft line breaks for whitespace fidelity testing."""
    body_parts = []

    body_parts.append(make_paragraph("Whitespace Fidelity Test", "Title"))
    body_parts.append(make_paragraph("Testing tabs and soft line breaks", "Subtitle"))

    body_parts.append(make_paragraph("Tab Characters", "Heading1"))

    # Paragraph with tabs
    body_parts.append(make_paragraph_with_whitespace([
        {'type': 'text', 'value': 'Column1'},
        {'type': 'tab'},
        {'type': 'text', 'value': 'Column2'},
        {'type': 'tab'},
        {'type': 'text', 'value': 'Column3'},
    ]))

    body_parts.append(make_paragraph_with_whitespace([
        {'type': 'text', 'value': 'Value A'},
        {'type': 'tab'},
        {'type': 'text', 'value': 'Value B'},
        {'type': 'tab'},
        {'type': 'text', 'value': 'Value C'},
    ]))

    body_parts.append(make_paragraph("Soft Line Breaks", "Heading1"))

    # Paragraph with soft line breaks (not paragraph breaks)
    body_parts.append(make_paragraph_with_whitespace([
        {'type': 'text', 'value': 'Line one of the paragraph'},
        {'type': 'br'},
        {'type': 'text', 'value': 'Line two (soft break)'},
        {'type': 'br'},
        {'type': 'text', 'value': 'Line three (soft break)'},
    ]))

    body_parts.append(make_paragraph("Mixed Whitespace", "Heading1"))

    # Paragraph with both tabs and line breaks
    body_parts.append(make_paragraph_with_whitespace([
        {'type': 'text', 'value': 'Name'},
        {'type': 'tab'},
        {'type': 'text', 'value': 'Age'},
        {'type': 'tab'},
        {'type': 'text', 'value': 'City'},
        {'type': 'br'},
        {'type': 'text', 'value': 'Alice'},
        {'type': 'tab'},
        {'type': 'text', 'value': '30'},
        {'type': 'tab'},
        {'type': 'text', 'value': 'NYC'},
        {'type': 'br'},
        {'type': 'text', 'value': 'Bob'},
        {'type': 'tab'},
        {'type': 'text', 'value': '25'},
        {'type': 'tab'},
        {'type': 'text', 'value': 'LA'},
    ]))

    body_parts.append(make_paragraph("Address Block", "Heading2"))

    # Address with carriage returns
    body_parts.append(make_paragraph_with_whitespace([
        {'type': 'text', 'value': 'John Doe'},
        {'type': 'br'},
        {'type': 'text', 'value': '123 Main Street'},
        {'type': 'br'},
        {'type': 'text', 'value': 'Anytown, ST 12345'},
    ]))

    body_parts.append(make_paragraph("Verification", "Heading1"))
    body_parts.append(make_paragraph(
        "The paragraphs above should preserve tab and line break characters in the output.",
        "Normal"
    ))

    body = '\n'.join(body_parts)
    create_docx("whitespace.docx", body)

def main():
    # Create fixtures directory
    FIXTURES_DIR.mkdir(parents=True, exist_ok=True)

    print("Creating test fixtures...")
    print(f"Output directory: {FIXTURES_DIR}")
    print()

    create_all_headings_fixture()
    create_mixed_lists_fixture()
    create_tables_fixture()
    create_comprehensive_fixture()
    create_large_fixture()
    create_wide_table_fixture()
    create_empty_fixture()
    create_minimal_fixture()
    create_whitespace_fixture()

    print()
    print("Done! Created 9 test fixtures.")

if __name__ == '__main__':
    main()

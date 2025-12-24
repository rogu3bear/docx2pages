#!/usr/bin/env python3
"""
DOCX Parser for docx2pages
Extracts document structure (headings, paragraphs, lists, tables) from DOCX files.
Outputs a JSON block list preserving document order.
"""

import json
import sys
import argparse
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path

# Word XML namespaces
NAMESPACES = {
    'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
}

def get_text(elem):
    """
    Extract all text from an element and its descendants.
    Preserves:
    - w:t text content
    - w:tab as \t
    - w:br (soft line break, no type) as \n
    - w:cr (carriage return) as \n
    """
    texts = []

    # Iterate through all children in document order
    for child in elem.iter():
        tag = child.tag.replace('{%s}' % NAMESPACES['w'], '')

        if tag == 't':
            # Text content
            if child.text:
                texts.append(child.text)
        elif tag == 'tab':
            # Tab character
            texts.append('\t')
        elif tag == 'br':
            # Break - check type
            br_type = child.get('{%s}type' % NAMESPACES['w'])
            if br_type is None:
                # Soft line break (no type attribute) -> newline
                texts.append('\n')
            # Page breaks (type="page") are handled separately, not as text
        elif tag == 'cr':
            # Carriage return
            texts.append('\n')

    return ''.join(texts)

def count_breaks_in_paragraph(para):
    """Count page breaks and section breaks in a paragraph."""
    breaks = 0

    # Page breaks: w:br with w:type="page"
    for br in para.iter('{%s}br' % NAMESPACES['w']):
        br_type = br.get('{%s}type' % NAMESPACES['w'])
        if br_type == 'page':
            breaks += 1

    # Section breaks in paragraph properties
    pPr = para.find('w:pPr', NAMESPACES)
    if pPr is not None:
        sectPr = pPr.find('w:sectPr', NAMESPACES)
        if sectPr is not None:
            breaks += 1

    return breaks

def parse_styles(styles_xml):
    """Parse styles.xml to build style hierarchy and detect heading styles."""
    style_info = {}

    if styles_xml is None:
        return style_info

    root = ET.fromstring(styles_xml)

    for style in root.findall('.//w:style', NAMESPACES):
        style_id = style.get('{%s}styleId' % NAMESPACES['w'])
        style_type = style.get('{%s}type' % NAMESPACES['w'])

        if style_type != 'paragraph':
            continue

        name_elem = style.find('w:name', NAMESPACES)
        name = name_elem.get('{%s}val' % NAMESPACES['w']) if name_elem is not None else style_id

        based_on_elem = style.find('w:basedOn', NAMESPACES)
        based_on = based_on_elem.get('{%s}val' % NAMESPACES['w']) if based_on_elem is not None else None

        style_info[style_id] = {
            'name': name,
            'based_on': based_on,
            'heading_level': None
        }

    # Detect heading levels from style names
    for style_id, info in style_info.items():
        name = info['name'].lower().replace(' ', '')

        if name == 'title':
            info['heading_level'] = 0
            info['is_title'] = True
        elif name == 'subtitle':
            info['heading_level'] = 0
            info['is_subtitle'] = True
        else:
            # Check for "Heading N" or "HeadingN" patterns
            for pattern in ['heading', 'heading ']:
                if name.startswith(pattern):
                    suffix = name[len(pattern):]
                    if suffix.isdigit():
                        info['heading_level'] = int(suffix)
                        break

    # Walk base_style chain for custom styles derived from headings
    def get_heading_level(style_id, visited=None):
        if visited is None:
            visited = set()
        if style_id in visited or style_id not in style_info:
            return None
        visited.add(style_id)

        info = style_info[style_id]
        if info['heading_level'] is not None:
            return info['heading_level']

        if info['based_on']:
            return get_heading_level(info['based_on'], visited)
        return None

    for style_id in style_info:
        if style_info[style_id]['heading_level'] is None:
            level = get_heading_level(style_id)
            if level is not None:
                style_info[style_id]['heading_level'] = level

    return style_info

def parse_numbering(numbering_xml):
    """Parse numbering.xml to understand list formats."""
    numbering_info = {
        'abstract_nums': {},
        'nums': {}
    }

    if numbering_xml is None:
        return numbering_info

    root = ET.fromstring(numbering_xml)

    # Parse abstract numbering definitions
    for abstract in root.findall('.//w:abstractNum', NAMESPACES):
        abstract_id = abstract.get('{%s}abstractNumId' % NAMESPACES['w'])
        levels = {}

        for lvl in abstract.findall('w:lvl', NAMESPACES):
            lvl_id = lvl.get('{%s}ilvl' % NAMESPACES['w'])
            num_fmt_elem = lvl.find('w:numFmt', NAMESPACES)
            num_fmt = num_fmt_elem.get('{%s}val' % NAMESPACES['w']) if num_fmt_elem is not None else 'bullet'

            # Determine if ordered based on numFmt
            # bullet = unordered, anything else (decimal, lowerLetter, etc.) = ordered
            is_ordered = num_fmt not in ('bullet', 'none')

            levels[lvl_id] = {
                'num_fmt': num_fmt,
                'is_ordered': is_ordered
            }

        numbering_info['abstract_nums'][abstract_id] = levels

    # Parse num definitions (link numId to abstractNumId)
    for num in root.findall('.//w:num', NAMESPACES):
        num_id = num.get('{%s}numId' % NAMESPACES['w'])
        abstract_ref = num.find('w:abstractNumId', NAMESPACES)
        if abstract_ref is not None:
            abstract_id = abstract_ref.get('{%s}val' % NAMESPACES['w'])
            numbering_info['nums'][num_id] = abstract_id

    return numbering_info

def get_list_info(para, numbering_info):
    """Check if paragraph is a list item and get list properties."""
    num_pr = para.find('.//w:numPr', NAMESPACES)
    if num_pr is None:
        return None

    num_id_elem = num_pr.find('w:numId', NAMESPACES)
    ilvl_elem = num_pr.find('w:ilvl', NAMESPACES)

    if num_id_elem is None:
        return None

    num_id = num_id_elem.get('{%s}val' % NAMESPACES['w'])
    ilvl = ilvl_elem.get('{%s}val' % NAMESPACES['w']) if ilvl_elem is not None else '0'

    # Skip numId="0" which means no numbering
    if num_id == '0':
        return None

    # Determine if ordered
    is_ordered = False
    if num_id in numbering_info['nums']:
        abstract_id = numbering_info['nums'][num_id]
        if abstract_id in numbering_info['abstract_nums']:
            levels = numbering_info['abstract_nums'][abstract_id]
            if ilvl in levels:
                is_ordered = levels[ilvl]['is_ordered']

    return {
        'num_id': num_id,
        'ilvl': int(ilvl),
        'is_ordered': is_ordered
    }

def get_paragraph_style_id(para):
    """Get the style ID of a paragraph."""
    pPr = para.find('w:pPr', NAMESPACES)
    if pPr is None:
        return None
    pStyle = pPr.find('w:pStyle', NAMESPACES)
    if pStyle is None:
        return None
    return pStyle.get('{%s}val' % NAMESPACES['w'])

def parse_table(table_elem):
    """Parse a table element into rows of cell texts."""
    rows = []
    for tr in table_elem.findall('w:tr', NAMESPACES):
        row = []
        for tc in tr.findall('w:tc', NAMESPACES):
            # Get all paragraphs in cell
            cell_texts = []
            for p in tc.findall('w:p', NAMESPACES):
                text = get_text(p)
                if text:
                    cell_texts.append(text)
            row.append('\n'.join(cell_texts))
        rows.append(row)
    return rows

def parse_docx(docx_path, verbose=False, preserve_breaks=False):
    """
    Parse a DOCX file and return a list of blocks in document order.

    Handles edge cases:
    - Empty DOCX files (no body content)
    - DOCX files with no styles.xml
    - Malformed XML in DOCX
    """
    blocks = []
    stats = {
        'headings': {},
        'paragraphs': 0,
        'lists': {'bulleted': 0, 'numbered': 0},
        'tables': {'count': 0, 'max_rows': 0, 'max_cols': 0},
        'warnings': [],
        'dropped_breaks': 0
    }

    # Validate ZIP structure
    try:
        zf = zipfile.ZipFile(docx_path, 'r')
    except zipfile.BadZipFile:
        stats['warnings'].append('Invalid ZIP file structure')
        return {'blocks': blocks, 'stats': stats}
    except FileNotFoundError:
        stats['warnings'].append('File not found')
        return {'blocks': blocks, 'stats': stats}

    try:
        # Check for document.xml (required)
        if 'word/document.xml' not in zf.namelist():
            stats['warnings'].append('No word/document.xml found - invalid DOCX')
            zf.close()
            return {'blocks': blocks, 'stats': stats}

        # Read document.xml
        try:
            doc_xml = zf.read('word/document.xml')
        except Exception as e:
            stats['warnings'].append(f'Error reading document.xml: {e}')
            zf.close()
            return {'blocks': blocks, 'stats': stats}

        # Read styles.xml if present (optional)
        styles_xml = None
        if 'word/styles.xml' in zf.namelist():
            try:
                styles_xml = zf.read('word/styles.xml')
            except Exception as e:
                stats['warnings'].append(f'Error reading styles.xml: {e}')

        # Read numbering.xml if present (optional)
        numbering_xml = None
        if 'word/numbering.xml' in zf.namelist():
            try:
                numbering_xml = zf.read('word/numbering.xml')
            except Exception as e:
                stats['warnings'].append(f'Error reading numbering.xml: {e}')
    finally:
        zf.close()

    # Parse styles with error handling
    try:
        style_info = parse_styles(styles_xml)
    except ET.ParseError as e:
        stats['warnings'].append(f'Malformed styles.xml: {e}')
        style_info = {}

    # Parse numbering with error handling
    try:
        numbering_info = parse_numbering(numbering_xml)
    except ET.ParseError as e:
        stats['warnings'].append(f'Malformed numbering.xml: {e}')
        numbering_info = {'abstract_nums': {}, 'nums': {}}

    # Parse document.xml with error handling
    try:
        root = ET.fromstring(doc_xml)
    except ET.ParseError as e:
        stats['warnings'].append(f'Malformed document.xml: {e}')
        return {'blocks': blocks, 'stats': stats}

    body = root.find('.//w:body', NAMESPACES)

    if body is None:
        stats['warnings'].append('No body element found in document')
        return {'blocks': blocks, 'stats': stats}

    # Track current list for grouping
    current_list = None
    current_list_identity = None  # (num_id, is_ordered)

    def flush_list():
        nonlocal current_list, current_list_identity
        if current_list is not None:
            blocks.append(current_list)
            if current_list['ordered']:
                stats['lists']['numbered'] += 1
            else:
                stats['lists']['bulleted'] += 1
            current_list = None
            current_list_identity = None

    # Iterate body children in order
    for child in body:
        tag = child.tag.replace('{%s}' % NAMESPACES['w'], '')

        if tag == 'p':
            # Count breaks in this paragraph
            break_count = count_breaks_in_paragraph(child)

            if break_count > 0:
                if preserve_breaks:
                    # Insert break blocks before processing paragraph
                    flush_list()
                    for _ in range(break_count):
                        blocks.append({'type': 'break'})
                else:
                    stats['dropped_breaks'] += break_count

            # Paragraph
            text = get_text(child)
            style_id = get_paragraph_style_id(child)
            list_info = get_list_info(child, numbering_info)

            # Check if it's a list item
            if list_info is not None:
                list_identity = (list_info['num_id'], list_info['is_ordered'])

                # Check if we should continue current list or start new one
                if current_list_identity != list_identity:
                    flush_list()
                    current_list = {
                        'type': 'list',
                        'ordered': list_info['is_ordered'],
                        'items': []
                    }
                    current_list_identity = list_identity

                current_list['items'].append({
                    'text': text,
                    'level': list_info['ilvl']
                })
                continue

            # Not a list item, flush any pending list
            flush_list()

            # Check for heading
            heading_level = None
            is_title = False
            is_subtitle = False

            if style_id and style_id in style_info:
                info = style_info[style_id]
                heading_level = info.get('heading_level')
                is_title = info.get('is_title', False)
                is_subtitle = info.get('is_subtitle', False)

            if heading_level is not None:
                if is_title:
                    block = {'type': 'title', 'text': text}
                    stats['headings']['title'] = stats['headings'].get('title', 0) + 1
                elif is_subtitle:
                    block = {'type': 'subtitle', 'text': text}
                    stats['headings']['subtitle'] = stats['headings'].get('subtitle', 0) + 1
                else:
                    block = {'type': 'heading', 'level': heading_level, 'text': text}
                    level_key = f'level_{heading_level}'
                    stats['headings'][level_key] = stats['headings'].get(level_key, 0) + 1
            else:
                block = {'type': 'paragraph', 'text': text}
                stats['paragraphs'] += 1

            # Skip empty paragraphs (but keep them for spacing)
            if text.strip() or block['type'] in ('heading', 'title', 'subtitle'):
                blocks.append(block)

        elif tag == 'tbl':
            # Table - flush any pending list first
            flush_list()

            rows = parse_table(child)
            block = {'type': 'table', 'rows': rows}
            blocks.append(block)

            stats['tables']['count'] += 1
            if rows:
                stats['tables']['max_rows'] = max(stats['tables']['max_rows'], len(rows))
                stats['tables']['max_cols'] = max(stats['tables']['max_cols'], max(len(r) for r in rows) if rows else 0)

        elif tag == 'sectPr':
            # Section properties at body level (final section)
            # These contain section breaks
            if preserve_breaks:
                flush_list()
                blocks.append({'type': 'break'})
            else:
                stats['dropped_breaks'] += 1

    # Flush any remaining list
    flush_list()

    # Add warning if breaks were dropped
    if stats['dropped_breaks'] > 0 and not preserve_breaks:
        stats['warnings'].append(f"Dropped {stats['dropped_breaks']} page/section break(s). Use --preserve-breaks to convert to blank paragraphs.")

    return {'blocks': blocks, 'stats': stats}

def main():
    parser = argparse.ArgumentParser(description='Parse DOCX file to JSON block list')
    parser.add_argument('input', help='Path to input DOCX file')
    parser.add_argument('--output', '-o', help='Output JSON file (default: stdout)')
    parser.add_argument('--verbose', '-v', action='store_true', help='Verbose output')
    parser.add_argument('--preserve-breaks', action='store_true',
                        help='Convert page/section breaks to empty paragraph blocks')

    args = parser.parse_args()

    if not Path(args.input).exists():
        print(f"Error: Input file not found: {args.input}", file=sys.stderr)
        sys.exit(1)

    result = parse_docx(args.input, verbose=args.verbose, preserve_breaks=args.preserve_breaks)

    output = json.dumps(result, indent=2, ensure_ascii=False)

    if args.output:
        with open(args.output, 'w', encoding='utf-8') as f:
            f.write(output)
        if args.verbose:
            print(f"Written to {args.output}", file=sys.stderr)
    else:
        print(output)

if __name__ == '__main__':
    main()

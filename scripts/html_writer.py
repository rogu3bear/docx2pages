#!/usr/bin/env python3
"""
HTML Writer for docx2pages
Converts JSON blocks to HTML for Pages import on macOS 26+.
"""

import json
import sys
from html import escape


def json_to_html(json_path: str, output_path: str) -> None:
    """
    Convert JSON blocks to HTML.

    Args:
        json_path: Path to JSON file from parse_docx.py
        output_path: Path for output HTML file
    """
    with open(json_path) as f:
        data = json.load(f)

    html = ['<!DOCTYPE html>', '<html>', '<head>',
            '<meta charset="utf-8">',
            '<style>',
            'body { font-family: -apple-system, Helvetica, Arial, sans-serif; }',
            'h1 { font-size: 24pt; }',
            'h2 { font-size: 18pt; }',
            'h3 { font-size: 14pt; }',
            'p { font-size: 12pt; }',
            'table { border-collapse: collapse; }',
            'td, th { border: 1px solid #ccc; padding: 8px; }',
            '</style>',
            '</head>', '<body>']

    for block in data.get('blocks', []):
        block_type = block.get('type')

        if block_type == 'title':
            text = escape(block.get('text', ''))
            html.append(f'<h1 style="font-size: 28pt;">{text}</h1>')

        elif block_type == 'subtitle':
            text = escape(block.get('text', ''))
            html.append(f'<h2 style="font-size: 18pt; color: #666;">{text}</h2>')

        elif block_type == 'heading':
            level = min(block.get('level', 1), 6)
            text = escape(block.get('text', ''))
            html.append(f'<h{level}>{text}</h{level}>')

        elif block_type == 'paragraph':
            text = escape(block.get('text', ''))
            # Preserve newlines and tabs
            text = text.replace('\n', '<br>').replace('\t', '&emsp;')
            html.append(f'<p>{text}</p>')

        elif block_type == 'list':
            ordered = block.get('ordered', False)
            tag = 'ol' if ordered else 'ul'
            html.append(f'<{tag}>')
            for item in block.get('items', []):
                item_text = escape(item.get('text', ''))
                indent = item.get('level', 0)
                style = f'margin-left: {indent * 20}px;' if indent > 0 else ''
                html.append(f'<li style="{style}">{item_text}</li>')
            html.append(f'</{tag}>')

        elif block_type == 'table':
            html.append('<table>')
            for row in block.get('rows', []):
                html.append('<tr>')
                for cell in row:
                    # Handle null cells (merged cell continuations)
                    cell_text = escape(cell) if cell is not None else ''
                    html.append(f'<td>{cell_text}</td>')
                html.append('</tr>')
            html.append('</table>')

        elif block_type == 'break':
            html.append('<br><br>')

    html.append('</body>')
    html.append('</html>')

    with open(output_path, 'w', encoding='utf-8') as f:
        f.write('\n'.join(html))


def main():
    if len(sys.argv) != 3:
        print(f"Usage: {sys.argv[0]} <input.json> <output.html>", file=sys.stderr)
        sys.exit(1)

    json_path = sys.argv[1]
    output_path = sys.argv[2]

    try:
        json_to_html(json_path, output_path)
        print(f"HTML written to: {output_path}")
    except Exception as e:
        print(f"Error: {e}", file=sys.stderr)
        sys.exit(1)


if __name__ == '__main__':
    main()

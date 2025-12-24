# Security Policy

## Supported Versions

| Version | Supported          |
| ------- | ------------------ |
| 1.2.x   | :white_check_mark: |
| < 1.2   | :x:                |

## Reporting a Vulnerability

If you discover a security vulnerability in docx2pages, please report it responsibly:

1. **Do not** open a public issue
2. Email the maintainers directly (see package.json or commit history for contact)
3. Include:
   - Description of the vulnerability
   - Steps to reproduce
   - Potential impact
   - Suggested fix (if any)

## Security Considerations

### File Handling

- docx2pages processes untrusted DOCX files
- The parser uses Python's stdlib `zipfile` and `xml.etree.ElementTree`
- Malformed DOCX files should fail gracefully, not crash or execute code

### Automation Permissions

- The tool requires macOS Automation permissions for Pages
- It does not request or use any other system permissions
- Template files are copied, never modified in place

### Lock File

- A lock file is created at `/tmp/docx2pages.lock`
- This is world-readable but only affects this tool

## Known Limitations

- No XML external entity (XXE) protection beyond Python's defaults
- Large malformed files may cause high memory usage before failing

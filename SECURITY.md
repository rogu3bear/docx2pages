# Security Policy

## Supported Versions

| Version | Supported          |
| ------- | ------------------ |
| 1.4.x   | :white_check_mark: |
| 1.3.x   | :x:                |
| < 1.3   | :x:                |

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
- The parser uses Python's `zipfile` and `defusedxml` for secure XML parsing
- XML External Entity (XXE) attacks are blocked by defusedxml
- Control characters invalid in JSON are sanitized from text content
- Malformed DOCX files should fail gracefully, not crash or execute code

### Automation Permissions

- The tool requires macOS Automation permissions for Pages
- It does not request or use any other system permissions
- Template files are copied, never modified in place

### Lock File

- A lock file is created at `/tmp/docx2pages.lock`
- This is world-readable but only affects this tool

## Known Limitations

- Large malformed files may cause high memory usage before failing
- Process timeout defaults to 120 seconds; very large documents may need `--timeout` adjustment

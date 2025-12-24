# Contributing to docx2pages

Thank you for your interest in contributing to docx2pages!

## Development Setup

1. Clone the repository
2. Ensure you have:
   - Swift 5.9+ toolchain
   - Python 3.x
   - Pages.app (for integration testing)

3. Build:
   ```bash
   swift build -c release
   ```

4. Run parser tests:
   ```bash
   python3 scripts/test_parser_golden.py
   ```

## Making Changes

1. Fork the repository
2. Create a feature branch: `git checkout -b feature/my-feature`
3. Make your changes
4. Run tests:
   ```bash
   # Parser golden tests (no Pages required)
   python3 scripts/test_parser_golden.py

   # Full smoke test (requires Pages + template)
   TEMPLATE=/path/to/template.pages scripts/smoke_test.sh
   ```
5. Commit with a descriptive message
6. Push and open a Pull Request

## Code Guidelines

- **Swift**: Follow Swift API Design Guidelines
- **Python**: Keep stdlib-only (no external dependencies)
- **JXA**: Keep automation minimal, avoid UI scripting

## What We Accept

- Bug fixes with regression tests
- Performance improvements with benchmarks
- Documentation improvements
- New fixture files for edge cases

## What We Don't Accept

- Features that require external Python dependencies
- UI scripting or AppleScript-based automation
- Changes that break strict mode semantics
- Changes that open the original template in Pages

## Questions?

Open an issue for discussion before starting large changes.

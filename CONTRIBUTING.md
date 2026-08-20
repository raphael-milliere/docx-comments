# Contributing to docx-comments

Thank you for your interest in contributing to docx-comments!

## Development Setup

1. Clone the repository:
   ```bash
   git clone https://github.com/raphael-milliere/docx-comments.git
   cd docx-comments
   ```

2. Create a virtual environment and install dependencies:
   ```bash
   uv venv
   source .venv/bin/activate  # On Windows: .venv\Scripts\activate
   uv pip install -e ".[dev]"
   ```

3. Verify the setup by running tests:
   ```bash
   pytest
   ```

## Running Tests

```bash
# Full test suite
pytest

# With coverage
pytest --cov=docx_comments

# Specific test file
pytest tests/test_manager_basic.py -v

# Specific test by name
pytest -k "test_add_comment"
```

## Code Quality

Before submitting a PR, ensure your code passes all checks:

```bash
# Linting
ruff check src/ tests/

# Formatting (check)
ruff format --check src/ tests/

# Formatting (fix)
ruff format src/ tests/

# Type checking
mypy src/docx_comments
```

## Submitting Changes

1. Fork the repository
2. Create a feature branch (`git checkout -b feature/your-feature`)
3. Make your changes
4. Run tests and linting
5. Commit your changes with a clear message
6. Push to your fork
7. Open a Pull Request

## Pull Request Guidelines

- Keep PRs focused on a single change
- Include tests for new functionality
- Update documentation if needed
- Ensure CI passes before requesting review

## Reporting Issues

When reporting bugs, please include:

- Python version
- python-docx version
- Minimal code example to reproduce the issue
- Expected vs actual behaviour
- Any error messages or tracebacks

## OOXML Resources

If you're working on comment-related features, these references are helpful:

- [ECMA-376 Standard](https://ecma-international.org/publications-and-standards/standards/ecma-376/) - Office Open XML specification
- [MS-DOCX Extensions](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-docx/) - Microsoft's DOCX extensions
- [CommentEx Class](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.office2013.word.commentex) - Threading documentation

## Questions?

Feel free to open an issue for questions or discussions about potential features.

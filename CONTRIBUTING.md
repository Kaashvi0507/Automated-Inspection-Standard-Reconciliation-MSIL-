# Contributing to GSIS-P / SMIR Reconciliation

Thank you for your interest in contributing! This document outlines our contribution process and expectations.

## Getting Started

1. **Fork** the repository
2. **Clone** your fork: `git clone https://github.com/YOUR-USERNAME/gsis-reconciliation.git`
3. **Create a virtual environment**:
   ```bash
   python -m venv venv
   source venv/bin/activate  # macOS/Linux
   # or
   venv\Scripts\activate  # Windows
   ```
4. **Install dependencies**:
   ```bash
   pip install -r requirements.txt
   pip install black pytest pytest-cov  # Dev dependencies
   ```

## Development Workflow

### Before You Start
- Check existing **Issues** & **Pull Requests** to avoid duplicates
- Open an Issue first for major features (get feedback before coding)
- Claim an issue by commenting — to avoid overlapping work

### Making Changes

1. **Create a feature branch**:
   ```bash
   git checkout -b feature/your-feature-name
   # or
   git checkout -b fix/bug-description
   git checkout -b docs/improvement
   ```

2. **Code**:
   - Follow **PEP 8** style guidelines
   - Add type hints where possible (Python 3.8+)
   - Write docstrings for functions (Google style)
   - Keep functions focused and testable

3. **Test Your Changes**:
   ```bash
   # Run the app locally to verify UI/UX changes
   streamlit run reconcile.py

   # Test critical functions in isolation (if applicable)
   pytest
   ```

4. **Format Code**:
   ```bash
   black reconcile.py
   ```

5. **Commit with clear messages**:
   ```bash
   git commit -m "feat: Add manual alignment override for unmatched rows"
   git commit -m "fix: Handle NaTType in date validation"
   git commit -m "docs: Update troubleshooting guide"
   git commit -m "refactor: Extract table parsing into separate module"
   ```
   Follow [Conventional Commits](https://www.conventionalcommits.org/):
   - `feat:` New feature
   - `fix:` Bug fix
   - `docs:` Documentation
   - `refactor:` Code restructuring (no behavior change)
   - `perf:` Performance improvement
   - `test:` Adding tests
   - `chore:` Build, CI, dependencies

6. **Push to your fork**:
   ```bash
   git push origin feature/your-feature-name
   ```

7. **Open a Pull Request** on the main repo:
   - Write a clear title and description
   - Link related issues: `Closes #123`
   - Include:
     - What changed and why
     - How to test the changes
     - Screenshots/GIFs if UI changes
     - Any breaking changes

### Pull Request Checklist
- [ ] Code follows style guidelines (Black formatted)
- [ ] All existing tests pass
- [ ] New functions have docstrings
- [ ] Type hints added (Python 3.8+)
- [ ] No hardcoded paths or credentials
- [ ] Commit messages follow conventional format
- [ ] No unnecessary dependencies added

## Code Style Guide

### Python Style
```python
"""Module docstring: one-liner describing purpose."""

from typing import Dict, List, Optional
import pandas as pd


def extract_metadata_from_pdf(pdf_bytes: bytes) -> Dict[str, str]:
    """Extract metadata from PDF header.
    
    Args:
        pdf_bytes: Raw PDF file content.
        
    Returns:
        Dictionary with keys: vendor_code, part_number, model_no, issue_date.
        
    Raises:
        ValueError: If PDF lacks required metadata.
    """
    # Implementation here
    pass


class ReconciliationEngine:
    """Manages PDF-to-Excel reconciliation workflow."""
    
    def __init__(self, schema: List[str]):
        """Initialize the engine.
        
        Args:
            schema: List of column names (19-column GSIS-P format).
        """
        self.schema = schema
```

### Naming Conventions
- `snake_case` for functions/variables: `extract_metadata()`, `confidence_score`
- `PascalCase` for classes: `ReconciliationEngine`, `PDFExtractor`
- `UPPER_SNAKE_CASE` for constants: `SCHEMA`, `MIC_PATTERN`, `NUMERIC_COLS`
- Avoid single-letter names except in loops: `for i in range(len(work))`

### Comments & Docstrings
```python
# High-level explanation of complex logic (above code)
# Explain the "why," not the "what"
confidence_score = 100  # Start at 100, deduct for risk factors

def validate_date(date_str: str) -> bool:
    """Check if date follows DD-MM-YYYY format."""
    return bool(re.match(r'^\d{2}-\d{2}-\d{4}$', date_str))
```

## Project Structure

```
gsis-reconciliation/
├── reconcile.py                  # Main Streamlit app
├── requirements.txt              # Dependencies
├── start.bat                     # Windows launcher
├── README.md                     # User guide
├── CONTRIBUTING.md               # This file
├── LICENSE                       # MIT License
├── .gitignore                    # Git ignore rules
├── .github/
│   ├── workflows/
│   │   ├── tests.yml             # GitHub Actions (tests)
│   │   └── lint.yml              # Code quality checks
│   └── ISSUE_TEMPLATE/
│       ├── bug_report.md
│       └── feature_request.md
└── docs/
    ├── architecture.md           # System design
    ├── api.md                    # Function reference
    └── troubleshooting.md        # Known issues & fixes
```

## Reporting Issues

### Bug Report
Use the [bug report template](.github/ISSUE_TEMPLATE/bug_report.md):
- **Title**: Short, descriptive
- **Steps to reproduce**: Exact actions leading to the bug
- **Expected behavior**: What should happen
- **Actual behavior**: What actually happens
- **Environment**: OS, Python version, key library versions
- **Attachments**: Error log, sample PDF/Excel (if shareable)

### Feature Request
Use the [feature request template](.github/ISSUE_TEMPLATE/feature_request.md):
- **Problem**: What user problem does this solve?
- **Proposed solution**: Your suggested implementation
- **Alternatives**: Other approaches you considered
- **Additional context**: Links, screenshots, etc.

## Review Process

1. **Automated checks**: GitHub Actions runs linting & tests
2. **Code review**: Maintainer reviews for:
   - Code quality & style
   - Functionality correctness
   - Documentation completeness
   - Performance impact
   - Security/safety
3. **Request changes** or **Approve** → Merge into `main`

## Release Process

Releases follow **Semantic Versioning** (`MAJOR.MINOR.PATCH`):
- `1.0.0` → Initial stable release
- `1.1.0` → New feature (backward-compatible)
- `1.0.1` → Bug fix (backward-compatible)
- `2.0.0` → Breaking changes

## Community Guidelines

- **Be respectful**: Treat all contributors with courtesy
- **Assume good intent**: Misunderstandings happen — clarify kindly
- **Help others**: Answer questions, review PRs, mentor newcomers
- **Share knowledge**: Document learnings, share insights

## Questions?

- **Discussions**: Use GitHub Discussions for questions
- **Issues**: Search existing issues before opening a new one
- **Email**: Reach out to the maintainer if needed

---

**Thank you for contributing to making GSIS-P reconciliation better!** 🙌

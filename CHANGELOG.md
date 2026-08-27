# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

### Added
- (Future features go here)

### Changed
- (Breaking changes go here)

### Deprecated
- (Soon-to-be-removed features go here)

### Removed
- (Removed features go here)

### Fixed
- (Bug fixes go here)

### Security
- (Security fixes go here)

---

## [1.0.0] - 2024-XX-XX

### Added
- Initial public release
- Three-phase reconciliation pipeline (metadata extraction, OCR, matching)
- Streamlit web UI with real-time editing
- Confidence scoring for data quality
- Fuzzy item matching with semantic embeddings
- Multi-engine OCR fallback (RapidOCR → pdfplumber → Docling)
- 19-column GSIS-P schema validation
- Export formats: Corrected Excel, Rebuilt-from-PDF, Diff CSV
- Manual alignment override for unmatched rows
- Batch renumbering and reordering by MIC
- PDF viewer with zoom and page navigation
- Searchable & filterable comparison grid
- LAN access via IP:8502

### Changed
- (N/A for initial release)

### Fixed
- Fixed `NaTType does not support strftime` crash in date validation
- Fixed Streamlit file watcher crash on PyTorch introspection
- Suppressed noisy HuggingFace offline warnings

### Security
- All models bundled locally (no internet at runtime)
- No credentials or sensitive data transmission

---

## Versioning Guidelines

### Major (X.0.0)
- Breaking changes to API or Excel schema
- Major refactoring
- New major features (e.g., multiple PDF matching)

### Minor (X.Y.0)
- New features (backward-compatible)
- Deprecation notices
- Performance improvements

### Patch (X.Y.Z)
- Bug fixes
- Security fixes
- Documentation updates
- Dependency updates (same minor version)

---

## Release Checklist

- [ ] Update version in `pyproject.toml`, `README.md`
- [ ] Update `CHANGELOG.md` with all changes
- [ ] Run full test suite: `pytest --cov`
- [ ] Format code: `black reconcile.py`
- [ ] Test on Windows, macOS, Linux (or CI passes)
- [ ] Create GitHub release with tag `vX.Y.Z`
- [ ] Add release notes to GitHub release
- [ ] Update documentation if needed

---

## Commit Message Format

Use [Conventional Commits](https://www.conventionalcommits.org/):

```
<type>(<scope>): <subject>

<body>

<footer>
```

### Types
- `feat`: A new feature
- `fix`: A bug fix
- `docs`: Documentation changes
- `style`: Code style (formatting, semicolons, etc.)
- `refactor`: Code refactoring without feature changes
- `perf`: Performance improvements
- `test`: Adding or updating tests
- `chore`: Build, CI, dependencies

### Examples
```
feat(ocr): Add Docling fallback for heavy PDF tables
fix(validation): Handle NaTType in date validation
docs: Update troubleshooting guide
refactor(matching): Extract fuzzy matching into separate module
perf(emoji): Optimize Streamlit re-renders
```

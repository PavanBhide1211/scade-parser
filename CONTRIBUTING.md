# Contributing to SCADE Model Parser

Thank you for taking the time to contribute. This document describes how to
report bugs, propose features, and submit pull requests.

---

## Table of Contents

- [Code of Conduct](#code-of-conduct)
- [Reporting Bugs](#reporting-bugs)
- [Suggesting Features](#suggesting-features)
- [Development Setup](#development-setup)
- [Submitting a Pull Request](#submitting-a-pull-request)
- [Coding Standards](#coding-standards)
- [Commit Message Format](#commit-message-format)

---

## Code of Conduct

Be respectful and constructive in all interactions. Harassment, personal attacks,
or discriminatory language will not be tolerated.

---

## Reporting Bugs

1. Search [existing issues](https://github.com/pavanbhide/scade-parser/issues)
   first to avoid duplicates.
2. Open a new issue and include:
   - Python version (`python --version`)
   - Operating system
   - The command you ran
   - The full output (stdout + stderr)
   - A minimal `.scade` file that reproduces the problem (if applicable and safe
     to share — strip any proprietary model content).

**Do not open a public issue for security vulnerabilities.**
See [SECURITY.md](SECURITY.md) instead.

---

## Suggesting Features

Open a GitHub issue with the label `enhancement`. Describe:

- The use case you are trying to solve.
- What the expected input and output would look like.
- Any SCADE XML structure examples that are relevant.

---

## Development Setup

```bash
# Clone the repo
git clone https://github.com/pavanbhide/scade-parser.git
cd scade-parser

# Create and activate a virtual environment
python -m venv .venv
source .venv/bin/activate       # Linux / macOS
.venv\Scripts\activate          # Windows

# Install runtime + development dependencies
pip install -r requirements.txt
pip install pytest ruff
```

### Running the smoke test

```bash
python scade_parser.py sample_model.scade --output /tmp/test_output.xlsx
```

Expected: exit code 0, 7 input signals, 6 output signals, 8 calibration parameters.

### Running the linter

```bash
ruff check scade_parser.py
```

---

## Submitting a Pull Request

1. Fork the repository and create a branch from `main`:
   ```bash
   git checkout -b fix/my-bug-description
   ```
2. Make your changes. Keep each PR focused on a single concern.
3. Update `CHANGELOG.md` under `[Unreleased]` with a brief description of your change.
4. Ensure `ruff check scade_parser.py` reports no errors.
5. Run the smoke test and confirm it passes.
6. Push your branch and open a pull request against `main`.
7. Fill in the PR template — describe _what_ changed and _why_.

### What makes a good PR

- Single-purpose — one bug fix or one feature per PR.
- Includes a test case or example `.scade` file if adding a new XML pattern.
- Does not introduce new non-stdlib dependencies without prior discussion.
- Maintains the single-file design of `scade_parser.py`.

---

## Coding Standards

- **Python 3.10+** — type hints using the built-in generics (`list[dict]`, not `List[Dict]`).
- **No dynamic execution** — no `eval()`, `exec()`, or `shell=True`.
- **No new hard dependencies** — prefer stdlib; if a third-party library is truly
  necessary, open an issue to discuss it first.
- **Security-first** — any code that opens files or parses untrusted input must
  go through the existing validation helpers (`_validate_input_path`, `_attr`, etc.).
- **Docstrings** — public functions and the main class should have a one-line
  docstring. Inline comments for non-obvious logic.
- **Line length** — 100 characters max (ruff default).

---

## Commit Message Format

Use the [Conventional Commits](https://www.conventionalcommits.org/) style:

```
<type>(<scope>): <short summary>

[optional body]

[optional footer]
```

Common types: `fix`, `feat`, `docs`, `refactor`, `test`, `chore`.

Examples:
```
fix(parser): handle Constant elements with no child Value node
feat(output): add auto-filter to all data sheets
docs: update README quick-start example
```

---

## Questions?

Open a [GitHub Discussion](https://github.com/pavanbhide/scade-parser/discussions)
for anything that is not a bug or feature request.

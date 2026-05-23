# Security Policy

## Supported Versions

| Version | Supported |
|---------|-----------|
| 1.0.x   | Yes       |

Only the latest release receives security fixes.

---

## Security Design

The SCADE Model Parser is designed to process XML files from potentially untrusted
sources safely. Key protections:

| Threat                        | Mitigation                                                                                      |
|-------------------------------|-------------------------------------------------------------------------------------------------|
| XML External Entity (XXE)     | `defusedxml` blocks all DOCTYPE and entity declarations. A stdlib fallback provides partial protection when `defusedxml` is unavailable — **`defusedxml` must be installed in any CI/CD or shared environment**. |
| XML bomb (billion-laughs)     | `defusedxml` detects recursive entity expansion. File-size pre-check (50 MB cap) provides an additional layer. |
| Oversized / malicious files   | Files exceeding `MAX_FILE_SIZE_MB` (default: 50 MB) are rejected before XML parsing begins.     |
| Unsupported file types        | Only `.scade` and `.xscade` extensions are opened; all others are silently skipped.             |
| Output path traversal         | The output filename is sanitised to `[A-Za-z0-9_.-]` only; the parent directory must pre-exist; the `.xlsx` extension is enforced. |
| Dynamic code execution        | The codebase contains no `eval()`, `exec()`, `compile()`, or `subprocess(shell=True)` calls.   |

---

## Reporting a Vulnerability

**Please do not report security vulnerabilities through public GitHub issues.**

To report a vulnerability:

1. Email **bhidepavan@gmail.com** with the subject line:
   `[SECURITY] scade-parser: <brief description>`

2. Include in your report:
   - A description of the vulnerability and its potential impact.
   - Steps to reproduce (minimal example file or command if possible).
   - Python version, OS, and `defusedxml` version (`pip show defusedxml`).

3. You will receive an acknowledgement within **48 hours** and a resolution
   timeline within **7 days**.

4. Once a fix is available and released, you will be credited in the
   `CHANGELOG.md` and the GitHub release notes (unless you prefer anonymity).

---

## Disclosure Policy

- Vulnerabilities are fixed in a private branch and released as a patch version.
- A GitHub Security Advisory is published at the time of the fix release.
- Public disclosure is coordinated with the reporter — typically 90 days after
  the fix is available, or sooner if the reporter prefers.

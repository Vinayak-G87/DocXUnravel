# Contributing

Thanks for helping improve DocXUnravel. Bug reports, documentation fixes, tests, and code changes are welcome.

## Report an issue

Search existing issues first. For a bug, include the command or Python call, expected result, actual result, Python version, and a small reproducible DOCX if possible. Use a synthetic document, not a customer file. Feature requests should describe the problem and a sample workflow.

Do not report security vulnerabilities in a public issue. Follow [SECURITY.md](SECURITY.md) instead.

## Send a change

1. Fork the repository and work on a branch in your fork. If a change is large, open an issue to discuss its shape first.
2. Install the project locally with `python -m pip install -e '.[dev]'`. For a smaller installation, use `python -m pip install -e .` and add only the extras your tests need.
3. Add a focused test when you change behavior. Keep examples generic and small. Do not include private documents, credentials, generated bundles, `build/`, `dist/`, or `.venv/`.
4. Run `python -m unittest discover -s tests` and `git diff --check`. The visual test may skip if LibreOffice cannot render in your environment; mention that in your pull request.
5. Open a pull request against the default branch. Say what changed, why, and what you tested. A maintainer reviews and merges it after CI passes.

The project is licensed under [MIT](LICENSE). By submitting a contribution, you agree to release your contribution under that license and confirm you have the right to do so. Maintainers handle version numbers and PyPI releases; contributors do not need publishing credentials.

# Releasing DocXUnravel to PyPI

This is a release checklist, not a record of configured GitHub/PyPI settings. No publish is performed by these instructions. The package name is `docxunravel`; confirm it is available to the intended PyPI account before the first release. PyPI project names are normalized, and an uploaded name/version/file cannot be overwritten or reused; choose a new version for every release (including failed or withdrawn releases).

## Local verification

From the repository root, on Python 3.11 or newer:

```sh
python3 -m venv .venv
.venv/bin/python -m pip install -e '.[dev]' build twine
.venv/bin/python -m unittest discover -s tests
.venv/bin/python -m build --sdist --wheel --outdir dist
.venv/bin/python -m twine check --strict dist/*.tar.gz dist/*.whl
```

Build from a clean checkout or remove *only* stale build artifacts you intend to replace before building. Check the newly generated files, not older archives already in `dist/`. The wheel should contain the `docxunravel` Python package, `schemas/*.json`, entry-point metadata and license metadata; it should not contain tests, fixtures, local documents or credentials. The sdist should contain the source package, schemas, `pyproject.toml`, README and license; `MANIFEST.in` prunes `tests/`. Inspect both archives (`unzip -l dist/*.whl` and `tar -tzf dist/*.tar.gz`) for unexpected files, and install the wheel in a fresh environment to check imports and the `docxunravel` CLI. Review any generated artifacts before uploading; don't treat checked-in configuration alone as proof of their contents.

## GitHub and PyPI setup (maintainer action)

1. Secure the repository: require 2FA for maintainers (through the owning organization where applicable), protect the default branch using a branch protection rule or ruleset, require passing CI checks and at least one independent review before merge, and restrict bypasses and tag creation to trusted maintainers. Select required check names only after CI exists and has run. Optionally add `CODEOWNERS` and require code-owner review if there are designated owners; neither is assumed to exist. Consider GitHub immutable releases if supported and compatible with the release process. Review release-workflow changes especially carefully.
2. Review `.github/workflows/ci-release.yml` before use. It tests pushes and pull requests, then on a **published GitHub release** builds the release tag, checks that the tag (for example `v0.4.0`) matches `project.version` (`0.4.0`) and artifact metadata, and passes the checked distributions to the publishing job. Prereleases do not publish to PyPI. Require its CI test checks on the protected branch once they have run. Keep permissions minimal: `id-token: write` is needed by the publishing job for OIDC; other jobs use `contents: read`. Publish the checked artifacts, not a separate rebuild.
3. Create a GitHub Actions environment named **`pypi`**, with required reviewers (maintainers other than the release initiator where possible) and environment deployment restrictions limited to release tags. The publishing job must declare `environment: pypi` so approval gates the actual upload. Do not store a PyPI API token in GitHub secrets or use password-based publishing.
4. In PyPI project publishing settings, configure a **Trusted Publisher** with GitHub owner **`Vinayak-G87`**, repository **`DocXUnravel`**, workflow filename **`ci-release.yml`**, and environment **`pypi`**. These fields must match the actual workflow filename and job environment exactly. For a first release, use PyPI's pending publisher flow if the project does not yet exist, and confirm ownership/name availability first. OIDC provides short-lived credentials to the approved job; no PyPI token is needed.

## Release

Update `project.version` in `pyproject.toml` and prepare release notes before merging reviewed changes. After CI passes, run the local verification above and create/push a matching protected version tag from the reviewed commit. **Publish a non-prerelease GitHub release for that tag** to trigger the workflow; pushing a tag alone does not publish to PyPI. Approve the `pypi` deployment only after confirming the tag, commit, version, check results and artifacts. Verify the resulting PyPI page and install the published wheel in a clean environment. If an upload is wrong, never attempt to overwrite it: publish a corrected new version. Do not publish from a developer machine or expose long-lived upload tokens.

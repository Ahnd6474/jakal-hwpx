# Releasing `jakal-hwpx`

This project already has GitHub Actions workflows for build and publish. This file covers the local checks that should happen before pushing a PyPI release.

## Release checklist

Before tagging a release:

1. Update the version in:
   - `pyproject.toml`
   - `src/jakal_hwpx/__init__.py`
2. Update user-facing docs if the public API changed:
   - `README.md`
   - `HWPX_MODULE.md`
   - `docs/hwpx-document.md`
3. Re-run the focused tests for the changed surface.
4. Rebuild `dist/`.
5. Run `twine check`.

## Local packaging setup

```bash
python -m pip install --upgrade build twine tox
```

## Recommended validation flow

### 1. Fast document-model gate

```bash
python -m pytest tests/test_document_model.py tests/test_hancom_document.py -q
```

### 2. CI-compatible release gate

```bash
tox -e release
```

This gate is expected to:

- run the release-oriented test matrix
- rebuild `dist/`
- run `python -m twine check dist/*`
- verify version consistency
- verify required packaged files

### 3. Full Windows release gate

On the Windows release machine with Hancom installed:

```powershell
python scripts/check_release.py --profile release
```

That profile adds the Hancom-facing smoke and roundtrip checks.

## Build commands

If you only want to prepare the package artifacts locally:

```bash
python -m build
python -m twine check dist/*
```

## TestPyPI

After a clean local gate:

```bash
python -m twine upload --repository testpypi dist/*
```

Then verify installation:

```bash
python -m pip install --upgrade pip
python -m pip install --index-url https://test.pypi.org/simple/ --extra-index-url https://pypi.org/simple jakal-hwpx
python -c "from jakal_hwpx import HwpxDocument; print(HwpxDocument)"
```

## PyPI

For a manual upload:

```bash
python -m twine upload dist/*
```

Preferred release sequence:

1. Update version and docs.
2. Run targeted tests.
3. Run `tox -e release`.
4. Run `python scripts/check_release.py --profile release` on the Windows release machine.
5. Upload to TestPyPI.
6. Smoke-test install from TestPyPI.
7. Upload to PyPI or trigger the publish workflow.

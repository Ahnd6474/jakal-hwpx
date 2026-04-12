# Releasing `jakal-hwpx`

This project already includes GitHub Actions workflows for building and publishing packages. This guide captures the last manual checks so a PyPI release is repeatable and low-risk.

The release gate is now defined by [`STABILITY_CONTRACT.md`](./STABILITY_CONTRACT.md) and enforced by [`scripts/check_release.py`](./scripts/check_release.py).

## Before the first public release

1. Choose the project's license and add a top-level `LICENSE` file.
2. Review bundled third-party tools and sample documents, then update `THIRD_PARTY_NOTICES.md` if the redistribution scope changed.
3. Create the `jakal-hwpx` project on TestPyPI and PyPI, or confirm that you own the existing project.
4. Configure PyPI trusted publishing for this repository.
5. Add the `pypi` and `testpypi` GitHub environments if they are not already configured.

Without a real `LICENSE` file and a reviewed third-party notice, packaging still builds, but the release is not ready for public redistribution.

## Local validation

Install the release tools:

```bash
python -m pip install --upgrade build twine tox
```

Run the CI-compatible release gate:

```bash
tox -e release
```

That command:

- validates the stability contract and sample corpus
- runs the full pytest suite
- runs the HWPX, HWP, and bridge stability matrices
- rebuilds `dist/`
- runs `python -m twine check dist/*`
- verifies the package version is consistent
- checks the wheel and source distribution for required package files
- warns if a top-level `LICENSE` file is still missing

For the full Windows release gate with Hancom validation, run:

```powershell
python scripts/check_release.py --profile release
```

That profile adds:

- Hancom corpus smoke validation
- zero-dialog enforcement for warning/recovery popups
- rejection of Hancom-related skips

If you only want the lightweight packaging check used by CI:

```bash
tox -e pkg
```

## TestPyPI

Use the publish workflow manually:

1. Open GitHub Actions.
2. Run the `Publish` workflow.
3. Select `testpypi`.
4. Install from TestPyPI and smoke-test the package:

```bash
python -m pip install --upgrade pip
python -m pip install --index-url https://test.pypi.org/simple/ --extra-index-url https://pypi.org/simple jakal-hwpx
python -c "from jakal_hwpx import HwpxDocument; print(HwpxDocument)"
```

## PyPI

There are two supported release paths:

1. Create a GitHub Release. The `Publish` workflow will publish to PyPI automatically.
2. Run the `Publish` workflow manually and choose `pypi`.

For a manual upload outside GitHub Actions:

```bash
python scripts/check_release.py --profile ci
python -m twine upload dist/*
```

## Suggested release sequence

1. Update the version in `pyproject.toml` and `src/jakal_hwpx/__init__.py`.
2. Run `tox -e py312` or the full test matrix you want before release.
3. Run `tox -e release`.
4. On the Windows release machine, run `python scripts/check_release.py --profile release`.
5. Publish to TestPyPI.
6. Smoke-test installation from TestPyPI.
7. Create the GitHub Release or run the PyPI publish workflow manually.

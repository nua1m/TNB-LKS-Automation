# Release Process

## Release Contract

The updater expects a GitHub Release with:

- a semantic version tag such as `v0.1.0`
- a ZIP asset produced by `scripts/build_release.py`
- repository files laid out at the ZIP root, not nested under an extra top-level folder
- packaged Windows binaries built by `scripts/build_windows_bundle.py`

The current artifact name format is:

```text
TNB-LKS-Automation-v<version>.zip
```

Example:

```text
TNB-LKS-Automation-v0.1.0.zip
```

## Included in the Release ZIP

The build script includes:

- `Run LKS Automation.bat`
- `launcher.exe`
- `updater.exe`
- `processor.exe`
- `_internal/`
- `VERSION`
- `README.md`
- `LKS Template (M).xlsm`

The build script excludes runtime and local-machine state such as:

- `.git/`
- `.venv/`
- `dist/`
- `build/`
- `uploads/`
- `results/`
- `__pycache__/`
- `*.pyc`

## Local Release Build

Build the release ZIP locally:

```powershell
.venv\Scripts\python.exe scripts\build_release.py
```

Build only the packaged Windows bundle:

```powershell
.venv\Scripts\python.exe scripts\build_windows_bundle.py
```

Output:

- `dist/TNB-LKS-Automation-v<version>.zip`
- `dist/release_manifest.json`
- `build/release-root/`

## Publish Flow

1. Update `VERSION`
2. Commit the release changes
3. Create and push a Git tag matching the version in `VERSION`

Example:

```powershell
git tag v0.1.0
git push origin v0.1.0
```

4. GitHub Actions runs `.github/workflows/release.yml`
5. The workflow builds the ZIP and publishes the GitHub Release asset

## Notes

- `updater.py` checks the latest GitHub Release, not Git commits directly.
- If multiple release assets exist, the updater prefers a `.zip` asset.
- Keep the tag format aligned with `VERSION` so update prompts remain consistent.
- GitHub Actions installs both `requirements.txt` and `requirements-build.txt` before building the release artifact.

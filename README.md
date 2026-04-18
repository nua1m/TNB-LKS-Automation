# TNB Smart Meter Reporting Automation

## Overview

This project automates the preparation of *Laporan Kerja Selesai* (LKS) workbooks for TNB smart meter installation reporting.

The current app is a Windows desktop workflow built around:
- a small Tkinter launcher for file selection
- a Python processing pipeline for Excel transformation
- output generation into a formatted `.xlsm` workbook

The immediate product direction is internal operational use first: make the tool easy to launch, update, and use for non-technical colleagues.

## Current Capabilities

- ingest raw `.xls` and `.xlsx` source files
- convert legacy `.xls` inputs before processing
- clean and normalize LKS source data
- remove `TRAS` rows
- skip duplicate service orders already present in the template
- write claim and attachment rows into the target workbook
- inject image evidence into the output workbook
- detect missing image slots and highlight defective rows
- save a final `.xlsm` file based on the bundled template

## Current Desktop Flow

1. Launch `Run LKS Automation.bat`
2. The updater checks GitHub Releases for a newer version
3. If an update exists, the user can approve the update
4. The launcher opens and prompts for the input Excel file
5. The processing pipeline generates the LKS output workbook

## Milestone A Status

Milestone A is the internal distribution milestone. The current repo now includes the first update foundation:
- `Run LKS Automation.bat` uses a relative path instead of a machine-specific hardcoded path
- `Run LKS Automation.bat` creates a local `.venv` on first run and reuses it afterward
- `updater.py` checks GitHub Releases before launching the app
- `VERSION` stores the local app version shown in the launcher
- `launcher.py` displays the current app version and exposes a manual update check

What is still needed to complete Milestone A:
- validate the `.venv` bootstrap flow on the other laptop
- make Python version setup more predictable across internal machines
- later return to signed packaging for broader distribution

The release/update contract now uses:
- `VERSION` as the local version source
- `Run LKS Automation.bat` to bootstrap the local `.venv` when missing
- `scripts/build_release.py` to produce the release ZIP
- `.github/workflows/release.yml` to publish tagged releases to GitHub
- `docs/RELEASE_PROCESS.md` as the release operator guide

## Current BA

This section captures the current business analysis assumptions and scope.

### Current Users

- primary users are internal operators: you and your colleague
- users are assumed to be non-technical and should not need Git or terminal usage

### Current Problem Being Solved

- manual LKS preparation takes too long
- manual copy-paste and screenshot work creates avoidable reporting errors
- updates are hard to distribute if the app depends on source-code workflows

### Current Business Value

- faster report preparation
- lower chance of missing-photo and formatting errors
- easier handoff to another operator on a different laptop
- simpler support model through release-based updates instead of Git commands

### Current Product Scope

- internal desktop tool first
- release-based updating for non-technical users
- UI improvements after the update path is in place
- future external/vendor distribution only after internal workflow is stable

### Future Business Path

If the tool is later sold to other vendors, the app should move toward:
- packaged Windows installer/executable distribution
- vendor-specific configuration instead of hardcoded rules
- clearer licensing and support boundaries
- a stable release channel for customer updates

## Tech Stack

- Python 3.10+
- Tkinter for the current desktop launcher
- Pandas for data processing
- OpenPyXL and Xlwings for Excel handling
- Requests and BeautifulSoup4 for supporting data and image workflows

## Local Run

Install dependencies:

```powershell
pip install -r requirements.txt
```

Run the app:

```powershell
python updater.py --launch launcher.py
```

Or use:

```powershell
Run LKS Automation.bat
```

On first run, the batch launcher:

- looks for Python 3.11 first, then 3.10, then `python`
- creates `.venv` if it does not already exist
- installs `requirements.txt`
- reuses the same `.venv` on later runs
- reinstalls dependencies only when `requirements.txt` changes

Build the release ZIP:

```powershell
.venv\Scripts\python.exe scripts\build_release.py
```

## Notes

- The updater depends on GitHub Releases, not Git pulls on the user machine.
- Until releases are published, the updater stays quiet during normal launch and will report release status only when the user clicks `Check Updates`.
- Proprietary TNB data and customer-specific material should stay out of the repository.

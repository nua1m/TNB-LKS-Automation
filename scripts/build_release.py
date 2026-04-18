from __future__ import annotations

import argparse
import json
import zipfile
from pathlib import Path


PROJECT_ROOT = Path(__file__).resolve().parent.parent
DIST_DIR = PROJECT_ROOT / "dist"
VERSION_FILE = PROJECT_ROOT / "VERSION"
APP_NAME = "TNB-LKS-Automation"
MANIFEST_NAME = "release_manifest.json"

INCLUDE_FILES = [
    "VERSION",
    "README.md",
    "requirements.txt",
    "Run LKS Automation.bat",
    "launcher.py",
    "main.py",
    "config.py",
    "updater.py",
    "LKS Template (M).xlsm",
]

INCLUDE_DIRS = [
    "core",
    "ui",
]

EXCLUDE_DIR_NAMES = {
    "__pycache__",
    ".git",
    ".venv",
    "dist",
    "build",
    "uploads",
    "results",
}

EXCLUDE_SUFFIXES = {
    ".pyc",
}


def get_version() -> str:
    return VERSION_FILE.read_text(encoding="utf-8").strip()


def should_exclude(path: Path) -> bool:
    return any(part in EXCLUDE_DIR_NAMES for part in path.parts) or path.suffix in EXCLUDE_SUFFIXES


def iter_release_files() -> list[Path]:
    files: list[Path] = []

    for relative_name in INCLUDE_FILES:
        candidate = PROJECT_ROOT / relative_name
        if not candidate.exists():
            raise FileNotFoundError(f"Required release file not found: {relative_name}")
        files.append(candidate)

    for relative_dir in INCLUDE_DIRS:
        root = PROJECT_ROOT / relative_dir
        if not root.exists():
            continue
        for path in root.rglob("*"):
            if path.is_file() and not should_exclude(path.relative_to(PROJECT_ROOT)):
                files.append(path)

    # Preserve order while removing duplicates.
    unique_files: list[Path] = []
    seen: set[Path] = set()
    for path in files:
        relative = path.relative_to(PROJECT_ROOT)
        if relative not in seen:
            unique_files.append(path)
            seen.add(relative)
    return unique_files


def build_manifest(version: str, archive_name: str, files: list[Path]) -> dict:
    return {
        "app_name": APP_NAME,
        "version": version,
        "archive_name": archive_name,
        "entrypoint": "Run LKS Automation.bat",
        "launch_script": "launcher.py",
        "updater_script": "updater.py",
        "included_files": [str(path.relative_to(PROJECT_ROOT)).replace("\\", "/") for path in files],
    }


def build_zip(version: str) -> Path:
    DIST_DIR.mkdir(parents=True, exist_ok=True)
    archive_name = f"{APP_NAME}-v{version}.zip"
    archive_path = DIST_DIR / archive_name

    files = iter_release_files()
    manifest = build_manifest(version, archive_name, files)

    with zipfile.ZipFile(archive_path, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        for path in files:
            archive.write(path, arcname=path.relative_to(PROJECT_ROOT))
        archive.writestr(MANIFEST_NAME, json.dumps(manifest, indent=2))

    manifest_path = DIST_DIR / MANIFEST_NAME
    manifest_path.write_text(json.dumps(manifest, indent=2) + "\n", encoding="utf-8")
    return archive_path


def main() -> int:
    parser = argparse.ArgumentParser(description="Build the release ZIP consumed by the updater.")
    parser.add_argument("--version", default="", help="Override VERSION file value for the release artifact name.")
    args = parser.parse_args()

    version = args.version.strip() or get_version()
    archive_path = build_zip(version)
    print(f"Built release artifact: {archive_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

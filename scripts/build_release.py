from __future__ import annotations

import argparse
import json
import subprocess
import sys
import zipfile
from pathlib import Path


PROJECT_ROOT = Path(__file__).resolve().parent.parent
DIST_DIR = PROJECT_ROOT / "dist"
BUILD_DIR = PROJECT_ROOT / "build"
VERSION_FILE = PROJECT_ROOT / "VERSION"
APP_NAME = "TNB-LKS-Automation"
MANIFEST_NAME = "release_manifest.json"
RELEASE_ROOT = BUILD_DIR / "release-root"


def get_version() -> str:
    return VERSION_FILE.read_text(encoding="utf-8").strip()


def iter_release_files() -> list[Path]:
    if not RELEASE_ROOT.exists():
        raise FileNotFoundError(
            f"Packaged release root not found: {RELEASE_ROOT}. Run scripts/build_windows_bundle.py first."
        )

    return [path for path in RELEASE_ROOT.rglob("*") if path.is_file()]


def build_manifest(version: str, archive_name: str, files: list[Path]) -> dict:
    return {
        "app_name": APP_NAME,
        "version": version,
        "archive_name": archive_name,
        "entrypoint": "Run LKS Automation.bat",
        "launcher_binary": "launcher.exe",
        "updater_binary": "updater.exe",
        "processor_binary": "processor.exe",
        "included_files": [str(path.relative_to(RELEASE_ROOT)).replace("\\", "/") for path in files],
    }


def build_zip(version: str) -> Path:
    DIST_DIR.mkdir(parents=True, exist_ok=True)
    archive_name = f"{APP_NAME}-v{version}.zip"
    archive_path = DIST_DIR / archive_name

    files = iter_release_files()
    manifest = build_manifest(version, archive_name, files)

    with zipfile.ZipFile(archive_path, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        for path in files:
            archive.write(path, arcname=path.relative_to(RELEASE_ROOT))
        archive.writestr(MANIFEST_NAME, json.dumps(manifest, indent=2))

    manifest_path = DIST_DIR / MANIFEST_NAME
    manifest_path.write_text(json.dumps(manifest, indent=2) + "\n", encoding="utf-8")
    return archive_path


def main() -> int:
    parser = argparse.ArgumentParser(description="Build the release ZIP consumed by the updater.")
    parser.add_argument("--version", default="", help="Override VERSION file value for the release artifact name.")
    parser.add_argument(
        "--skip-package",
        action="store_true",
        help="Skip rebuilding the Windows bundle before creating the release ZIP.",
    )
    args = parser.parse_args()

    version = args.version.strip() or get_version()
    if not args.skip_package:
        subprocess.run([sys.executable, str(PROJECT_ROOT / "scripts" / "build_windows_bundle.py")], check=True)
    archive_path = build_zip(version)
    print(f"Built release artifact: {archive_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

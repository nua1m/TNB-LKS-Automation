from __future__ import annotations

import shutil
import subprocess
import sys
from pathlib import Path


PROJECT_ROOT = Path(__file__).resolve().parent.parent
BUILD_DIR = PROJECT_ROOT / "build"
PYINSTALLER_WORK_DIR = BUILD_DIR / "pyinstaller"
PYINSTALLER_DIST_DIR = BUILD_DIR / "pyinstaller-dist"
SPEC_DIR = BUILD_DIR / "spec"
RELEASE_ROOT = BUILD_DIR / "release-root"
VERSION = (PROJECT_ROOT / "VERSION").read_text(encoding="utf-8").strip()


def run_pyinstaller(entry_script: str, name: str, windowed: bool, hidden_imports: list[str] | None = None) -> None:
    command = [
        sys.executable,
        "-m",
        "PyInstaller",
        "--noconfirm",
        "--clean",
        "--onedir",
        "--name",
        name,
        "--distpath",
        str(PYINSTALLER_DIST_DIR),
        "--workpath",
        str(PYINSTALLER_WORK_DIR),
        "--specpath",
        str(SPEC_DIR),
    ]

    if windowed:
        command.append("--windowed")

    for hidden_import in hidden_imports or []:
        command.extend(["--hidden-import", hidden_import])

    command.append(str(PROJECT_ROOT / entry_script))
    subprocess.run(command, check=True, cwd=str(PROJECT_ROOT))


def prepare_release_root() -> Path:
    if RELEASE_ROOT.exists():
        shutil.rmtree(RELEASE_ROOT)
    RELEASE_ROOT.mkdir(parents=True, exist_ok=True)

    for app_name in ("launcher", "updater", "processor"):
        source_dir = PYINSTALLER_DIST_DIR / app_name
        if not source_dir.exists():
            raise FileNotFoundError(f"Missing packaged app directory: {source_dir}")

        for item in source_dir.iterdir():
            destination = RELEASE_ROOT / item.name
            if item.is_dir():
                if destination.exists():
                    shutil.rmtree(destination)
                shutil.copytree(item, destination)
            else:
                shutil.copy2(item, destination)

    for extra_file in (
        "Run LKS Automation.bat",
        "VERSION",
        "README.md",
        "LKS Template (M).xlsm",
    ):
        shutil.copy2(PROJECT_ROOT / extra_file, RELEASE_ROOT / extra_file)

    return RELEASE_ROOT


def main() -> int:
    for directory in (PYINSTALLER_WORK_DIR, PYINSTALLER_DIST_DIR, SPEC_DIR):
        directory.mkdir(parents=True, exist_ok=True)

    run_pyinstaller("launcher.py", "launcher", windowed=True)
    run_pyinstaller("updater.py", "updater", windowed=True)
    run_pyinstaller(
        "processor.py",
        "processor",
        windowed=False,
        hidden_imports=["win32com", "win32com.client", "pythoncom", "pywintypes"],
    )

    release_root = prepare_release_root()
    print(f"Prepared Windows bundle: {release_root}")
    print(f"Version: {VERSION}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

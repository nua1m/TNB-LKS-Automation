from __future__ import annotations

import argparse
import os
import shutil
import subprocess
import sys
import zipfile
from pathlib import Path
from tkinter import Tk, messagebox

import requests


REPOSITORY = "nua1m/TNB-LKS-Automation"
LATEST_RELEASE_API = f"https://api.github.com/repos/{REPOSITORY}/releases/latest"
REQUEST_TIMEOUT_SECONDS = 8
PRESERVE_TOP_LEVEL = {
    ".git",
    ".venv",
    "__pycache__",
    "Data",
    "results",
    "uploads",
}


def get_app_dir() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


APP_DIR = get_app_dir()
VERSION_FILE = APP_DIR / "VERSION"


def parse_version(raw: str) -> tuple[int, ...]:
    raw = raw.strip().lstrip("vV")
    if not raw:
        return (0,)

    parts = []
    for piece in raw.split("."):
        digits = "".join(ch for ch in piece if ch.isdigit())
        parts.append(int(digits) if digits else 0)
    return tuple(parts) if parts else (0,)


def get_local_version() -> str:
    if VERSION_FILE.exists():
        return VERSION_FILE.read_text(encoding="utf-8").strip()
    return "0.0.0"


def get_latest_release() -> dict | None:
    response = requests.get(
        LATEST_RELEASE_API,
        headers={"Accept": "application/vnd.github+json"},
        timeout=REQUEST_TIMEOUT_SECONDS,
    )
    if response.status_code == 404:
        return None
    response.raise_for_status()

    release = response.json()
    return {
        "version": release.get("tag_name", "").strip(),
        "name": release.get("name", "").strip() or release.get("tag_name", "").strip(),
        "notes": (release.get("body") or "").strip(),
        "zip_url": pick_release_zip_url(release),
    }


def pick_release_zip_url(release: dict) -> str:
    for asset in release.get("assets", []):
        asset_name = (asset.get("name") or "").lower()
        if asset_name.endswith(".zip") and asset.get("browser_download_url"):
            return asset["browser_download_url"]

    raise RuntimeError(
        "Latest release does not contain the packaged ZIP asset yet. "
        "Wait for the release workflow to finish and try again."
    )


def should_update(local_version: str, remote_version: str) -> bool:
    return parse_version(remote_version) > parse_version(local_version)


def download_release_zip(zip_url: str, destination: Path) -> None:
    with requests.get(zip_url, stream=True, timeout=REQUEST_TIMEOUT_SECONDS) as response:
        response.raise_for_status()
        with destination.open("wb") as file_handle:
            for chunk in response.iter_content(chunk_size=1024 * 128):
                if chunk:
                    file_handle.write(chunk)


def unpack_release(zip_path: Path, extract_dir: Path) -> Path:
    with zipfile.ZipFile(zip_path) as archive:
        archive.extractall(extract_dir)

    children = [path for path in extract_dir.iterdir() if path.is_dir()]
    if len(children) == 1:
        return children[0]
    return extract_dir


def should_skip(relative_path: Path) -> bool:
    first_part = relative_path.parts[0] if relative_path.parts else ""
    return first_part in PRESERVE_TOP_LEVEL


def apply_release(payload_dir: Path) -> None:
    for source_path in payload_dir.rglob("*"):
        if source_path.is_dir():
            continue

        relative_path = source_path.relative_to(payload_dir)
        if should_skip(relative_path):
            continue

        destination = APP_DIR / relative_path
        destination.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(source_path, destination)


def schedule_windows_update(payload_dir: Path, launch_target_name: str) -> None:
    temp_root = payload_dir.parent
    script_path = temp_root / "apply_update.cmd"
    launch_target_path = APP_DIR / launch_target_name

    script_lines = [
        "@echo off",
        "setlocal",
        "ping 127.0.0.1 -n 3 >nul",
        f'xcopy "{payload_dir}\\*" "{APP_DIR}\\" /E /I /Y >nul',
    ]

    if launch_target_name:
        script_lines.append(f'start "" "{launch_target_path}"')

    script_lines.extend(
        [
            f'rd /s /q "{temp_root}" >nul 2>nul',
            "exit /b 0",
        ]
    )

    script_path.write_text("\r\n".join(script_lines) + "\r\n", encoding="utf-8")
    creation_flags = 0
    if hasattr(subprocess, "DETACHED_PROCESS"):
        creation_flags |= subprocess.DETACHED_PROCESS
    if hasattr(subprocess, "CREATE_NEW_PROCESS_GROUP"):
        creation_flags |= subprocess.CREATE_NEW_PROCESS_GROUP

    subprocess.Popen(
        ["cmd", "/c", str(script_path)],
        cwd=str(APP_DIR),
        creationflags=creation_flags,
        close_fds=True,
    )


def show_update_prompt(release: dict, local_version: str) -> bool:
    title = "TNB LKS Automation Update"
    message = (
        f"Current version: {local_version}\n"
        f"Latest version: {release['version']}\n\n"
        f"{release['name']}\n\n"
        "Update now?"
    )
    return messagebox.askyesno(title, message)


def launch_target(target: str) -> int:
    target_path = APP_DIR / target
    if target_path.suffix.lower() == ".py":
        return subprocess.call([sys.executable, str(target_path)], cwd=str(APP_DIR))
    return subprocess.call([str(target_path)], cwd=str(APP_DIR))


def check_and_apply_updates(interactive: bool, show_status: bool, launch_target_name: str) -> bool:
    local_version = get_local_version()

    try:
        release = get_latest_release()
    except requests.RequestException as exc:
        if interactive and show_status:
            messagebox.showwarning(
                "Update Check Failed",
                f"Could not check for updates.\n\n{exc}",
            )
        return False
    except Exception as exc:
        if interactive and show_status:
            messagebox.showwarning(
                "Update Check Failed",
                f"Could not check for updates.\n\n{exc}",
            )
        return False

    if not release or not release["version"]:
        if interactive and show_status:
            messagebox.showinfo(
                "No Releases Yet",
                "No GitHub release has been published yet for this app.",
            )
        return False

    if not should_update(local_version, release["version"]):
        if interactive and show_status:
            messagebox.showinfo(
                "Up To Date",
                f"You are already on the latest version ({local_version}).",
            )
        return False

    if interactive and not show_update_prompt(release, local_version):
        return False

    temp_root = Path(os.environ.get("TEMP", APP_DIR)) / f"tnb_lks_update_{release['version'].replace('.', '_')}"
    try:
        if temp_root.exists():
            shutil.rmtree(temp_root)
        temp_root.mkdir(parents=True, exist_ok=True)

        archive_path = temp_root / "release.zip"
        payload_dir = temp_root / "payload"
        download_release_zip(release["zip_url"], archive_path)
        unpacked_dir = unpack_release(archive_path, payload_dir)

        if getattr(sys, "frozen", False) and sys.platform.startswith("win"):
            schedule_windows_update(unpacked_dir, launch_target_name)
        else:
            apply_release(unpacked_dir)
    except Exception as exc:
        if interactive:
            messagebox.showerror(
                "Update Failed",
                f"Could not apply the latest release.\n\n{exc}",
            )
        return False

    if interactive:
        messagebox.showinfo(
            "Update Complete",
            f"Updated to version {release['version']}.",
        )
    return True


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--launch", default="", help="Script to launch after the update check.")
    parser.add_argument(
        "--check-only",
        action="store_true",
        help="Check for updates but do not launch a target script.",
    )
    args = parser.parse_args()

    root = Tk()
    root.withdraw()

    updated = check_and_apply_updates(
        interactive=True,
        show_status=args.check_only,
        launch_target_name=args.launch,
    )

    if args.check_only:
        root.destroy()
        return 0

    root.destroy()

    if updated and getattr(sys, "frozen", False) and sys.platform.startswith("win"):
        return 0

    if args.launch:
        return launch_target(args.launch)

    return 0


if __name__ == "__main__":
    raise SystemExit(main())

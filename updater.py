from __future__ import annotations

import argparse
import hashlib
import shutil
import subprocess
import sys
import tempfile
import zipfile
from pathlib import Path
from tkinter import Tk, messagebox

import requests


APP_DIR = Path(__file__).resolve().parent
VERSION_FILE = APP_DIR / "VERSION"
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
REQ_MARKER = APP_DIR / ".venv" / "requirements.sha256"


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

    zipball_url = release.get("zipball_url")
    if not zipball_url:
        raise RuntimeError("Latest release does not contain a downloadable ZIP.")
    return zipball_url


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


def calculate_requirements_hash() -> str:
    requirements_path = APP_DIR / "requirements.txt"
    digest = hashlib.sha256(requirements_path.read_bytes()).hexdigest().upper()
    return digest


def sync_requirements() -> None:
    requirements_path = APP_DIR / "requirements.txt"
    if not requirements_path.exists():
        return

    current_hash = calculate_requirements_hash()
    installed_hash = REQ_MARKER.read_text(encoding="utf-8").strip() if REQ_MARKER.exists() else ""
    if current_hash == installed_hash:
        return

    subprocess.check_call([sys.executable, "-m", "pip", "install", "-r", str(requirements_path)], cwd=str(APP_DIR))
    REQ_MARKER.parent.mkdir(parents=True, exist_ok=True)
    REQ_MARKER.write_text(current_hash + "\n", encoding="utf-8")


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
    return subprocess.call([sys.executable, str(target_path)], cwd=str(APP_DIR))


def check_and_apply_updates(interactive: bool, show_status: bool) -> bool:
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

    try:
        with tempfile.TemporaryDirectory(prefix="tnb_lks_update_") as temp_dir:
            temp_path = Path(temp_dir)
            archive_path = temp_path / "release.zip"
            download_release_zip(release["zip_url"], archive_path)
            payload_dir = unpack_release(archive_path, temp_path / "payload")
            apply_release(payload_dir)
            sync_requirements()
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

    check_and_apply_updates(
        interactive=True,
        show_status=args.check_only,
    )

    if args.check_only:
        root.destroy()
        return 0

    root.destroy()

    if args.launch:
        return launch_target(args.launch)

    return 0


if __name__ == "__main__":
    raise SystemExit(main())

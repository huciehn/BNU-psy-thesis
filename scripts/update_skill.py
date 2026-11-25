"""Check GitHub Releases and install a tested update with a local backup."""
import argparse
import json
import os
from pathlib import Path, PurePosixPath
import re
import shutil
import subprocess
import sys
import tempfile
import uuid
import zipfile

import requests

ROOT = Path(__file__).resolve().parents[1]


def version_tuple(value):
    match = re.fullmatch(r"v?(\d+)\.(\d+)\.(\d+)", value.strip())
    if not match:
        raise ValueError("Versions must use MAJOR.MINOR.PATCH")
    return tuple(int(part) for part in match.groups())


def local_version():
    profile = json.loads((ROOT / "references/profile.json").read_text(encoding="utf-8"))
    return profile["version"]


def latest_release(repository):
    response = requests.get(
        f"https://api.github.com/repos/{repository}/releases/latest",
        headers={"Accept": "application/vnd.github+json", "User-Agent": "BNU-psy-thesis-update"},
        timeout=15,
    )
    response.raise_for_status()
    return response.json()


def asset_url(release, name):
    return next((item["browser_download_url"] for item in release.get("assets", []) if item.get("name") == name), None)


def download(url, path):
    response = requests.get(url, headers={"User-Agent": "BNU-psy-thesis-update"}, timeout=60)
    response.raise_for_status()
    path.write_bytes(response.content)


def validate_archive(archive, manifest_path, expected_root):
    import hashlib
    manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
    digest = hashlib.sha256(archive.read_bytes()).hexdigest()
    if digest != manifest.get("sha256"):
        raise ValueError("Release archive hash does not match manifest")
    with zipfile.ZipFile(archive) as package:
        if package.testzip() is not None:
            raise ValueError("Release archive is damaged")
        for name in package.namelist():
            pure = PurePosixPath(name)
            if pure.is_absolute() or ".." in pure.parts or not pure.parts or pure.parts[0] != expected_root:
                raise ValueError("Release archive contains an unsafe path")
            info = package.getinfo(name)
            if not info.is_dir():
                target = Path(pure.parts[-1])
                if not target.name:
                    raise ValueError("Invalid release file")
        package.extractall(archive.parent)


def install(staging_parent, expected_root):
    staged = staging_parent / expected_root
    skill = (staged / "SKILL.md").read_text(encoding="utf-8")
    if not re.search(r"^name: BNU-psy-thesis$", skill, re.M):
        raise ValueError("Unexpected skill identity")
    completed = subprocess.run(
        [sys.executable, "-m", "unittest", "discover", "-s", "tests", "-v"],
        cwd=staged,
        capture_output=True,
        text=True,
        timeout=180,
    )
    if completed.returncode:
        raise ValueError("Release tests failed: " + (completed.stderr or completed.stdout)[-2000:])
    backups = ROOT.parent / ".BNU-psy-thesis-backups"
    backups.mkdir(parents=True, exist_ok=True)
    backup = backups / (uuid.uuid4().hex + "-previous")
    os.replace(ROOT, backup)
    try:
        os.replace(staged, ROOT)
    except Exception:
        os.replace(backup, ROOT)
        raise
    for name in (".git", "_local", "dist"):
        preserved = backup / name
        if preserved.exists() and not (ROOT / name).exists():
            shutil.move(str(preserved), str(ROOT / name))
    return backup


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--auto", action="store_true", help="Install when a newer release exists")
    args = parser.parse_args()
    result = {"status": "unknown", "local_version": local_version(), "updated": False}
    try:
        config = json.loads((ROOT / "references/update.json").read_text(encoding="utf-8"))
        release = latest_release(config["repository"])
        latest = release["tag_name"]
        result["latest_version"] = latest
        if version_tuple(latest) <= version_tuple(result["local_version"]):
            result["status"] = "up_to_date"
        else:
            result["status"] = "update_available"
            if args.auto:
                asset_name = config.get("asset", "BNU-psy-thesis.skill")
                archive_url = asset_url(release, asset_name)
                manifest_url = asset_url(release, "manifest.json")
                if not archive_url or not manifest_url:
                    raise ValueError("Latest release lacks the skill archive or manifest")
                staging = Path(tempfile.mkdtemp(prefix="BNU-psy-thesis-update-", dir=ROOT.parent))
                try:
                    archive = staging / asset_name
                    manifest = staging / "manifest.json"
                    download(archive_url, archive)
                    download(manifest_url, manifest)
                    validate_archive(archive, manifest, "BNU-psy-thesis")
                    backup = install(staging, "BNU-psy-thesis")
                    result.update({"updated": True, "status": "updated", "backup": str(backup)})
                finally:
                    shutil.rmtree(staging, ignore_errors=True)
    except Exception as exc:
        result["status"] = "network_or_update_unavailable" if not args.auto else "update_failed"
        result["message"] = str(exc)
        if args.auto:
            print(json.dumps(result, ensure_ascii=False))
            return 0
    print(json.dumps(result, ensure_ascii=False))
    return 0


if __name__ == "__main__":
    sys.exit(main())

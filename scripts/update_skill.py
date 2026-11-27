"""Install a tested GitHub release without retaining local backup copies."""
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


def get_release(repository, version=None):
    suffix = "/releases/latest" if version is None else "/releases/tags/" + version
    response = requests.get(
        f"https://api.github.com/repos/{repository}{suffix}",
        headers={"Accept": "application/vnd.github+json", "User-Agent": "BNU-psy-thesis-update"},
        timeout=15,
    )
    response.raise_for_status()
    return response.json()


def latest_release(repository):
    return get_release(repository)


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
    root_files = {"SKILL.md", "README.md", "LICENSE", "requirements.txt", "requirements-optional.txt", ".gitignore"}
    trees = {"scripts": {".py"}, "references": {".md", ".json"}, "tests": {".py"}, "examples": {".json"}, "assets": {".doc", ".docx", ".pdf", ".md"}}
    def public_files(base):
        result = [base / name for name in root_files if (base / name).is_file()]
        for directory, suffixes in trees.items():
            for path in (base / directory).rglob("*"):
                relative = path.relative_to(base)
                if path.is_file() and path.suffix in suffixes and "__pycache__" not in relative.parts and "_local" not in relative.parts:
                    result.append(path)
        return result
    backup = staging_parent / (".transaction-" + uuid.uuid4().hex)
    backup.mkdir(parents=True)
    old_files = public_files(ROOT)
    for source in old_files:
        relative = source.relative_to(ROOT)
        target = backup / relative
        target.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(source, target)
    new_files = public_files(staged)
    new_relative = {path.relative_to(staged) for path in new_files}
    old_relative = {path.relative_to(ROOT) for path in old_files}
    try:
        for source in new_files:
            relative = source.relative_to(staged)
            target = ROOT / relative
            target.parent.mkdir(parents=True, exist_ok=True)
            temporary = target.with_name(target.name + ".update-" + uuid.uuid4().hex)
            shutil.copy2(source, temporary)
            os.replace(temporary, target)
        for old in old_files:
            relative = old.relative_to(ROOT)
            if relative not in new_relative and old.exists():
                old.unlink()
        archive = staging_parent / "BNU-psy-thesis.skill"
        manifest = staging_parent / "manifest.json"
        if archive.is_file() and manifest.is_file():
            if (ROOT / "dist").exists():
                shutil.move(str(ROOT / "dist"), str(backup / "dist"))
            (ROOT / "dist").mkdir()
            shutil.copy2(archive, ROOT / "dist" / archive.name)
            shutil.copy2(manifest, ROOT / "dist" / manifest.name)
    except Exception:
        for source in new_files:
            relative = source.relative_to(staged)
            target = ROOT / relative
            if relative not in old_relative and target.is_file():
                target.unlink()
        for source in backup.rglob("*"):
            if source.is_file():
                relative = source.relative_to(backup)
                target = ROOT / relative
                target.parent.mkdir(parents=True, exist_ok=True)
                shutil.copy2(source, target)
        raise
    finally:
        shutil.rmtree(backup, ignore_errors=True)


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    action = parser.add_mutually_exclusive_group()
    action.add_argument("--auto", action="store_true", help="Install when a newer release exists")
    action.add_argument("--version", help="Install a selected release, including an older version")
    args = parser.parse_args()
    result = {"status": "unknown", "local_version": local_version(), "updated": False}
    try:
        config = json.loads((ROOT / "references/update.json").read_text(encoding="utf-8"))
        if args.version:
            version_tuple(args.version)
        release = get_release(config["repository"], args.version)
        latest = release["tag_name"]
        result["latest_version"] = latest
        install_release = bool(args.version)
        if args.version:
            result["status"] = "selected_version"
        elif version_tuple(latest) <= version_tuple(result["local_version"]):
            result["status"] = "up_to_date"
        else:
            result["status"] = "update_available"
            install_release = args.auto
        if install_release:
            asset_name = config.get("asset", "BNU-psy-thesis.skill")
            archive_url = asset_url(release, asset_name)
            manifest_url = asset_url(release, "manifest.json")
            if not archive_url or not manifest_url:
                raise ValueError("Selected release lacks the skill archive or manifest")
            staging = Path(tempfile.mkdtemp(prefix="BNU-psy-thesis-update-", dir=ROOT.parent))
            try:
                archive = staging / asset_name
                manifest = staging / "manifest.json"
                download(archive_url, archive)
                download(manifest_url, manifest)
                validate_archive(archive, manifest, "BNU-psy-thesis")
                install(staging, "BNU-psy-thesis")
                result.update({"updated": True, "status": "updated"})
            finally:
                shutil.rmtree(staging, ignore_errors=True)
    except Exception as exc:
        result["status"] = "update_failed" if args.auto or args.version else "network_or_update_unavailable"
        result["message"] = str(exc)
        if args.auto:
            print(json.dumps(result, ensure_ascii=False))
            return 0
        if args.version:
            print(json.dumps(result, ensure_ascii=False))
            return 1
    print(json.dumps(result, ensure_ascii=False))
    return 0


if __name__ == "__main__":
    sys.exit(main())

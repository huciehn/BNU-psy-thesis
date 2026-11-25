"""Package a public, portable skill using an explicit allowlist."""
import argparse
import hashlib
import json
import re
import shutil
import uuid
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
ROOT_FILES = ("SKILL.md", "README.md", "LICENSE", "requirements.txt", "requirements-optional.txt", ".gitignore")
TREES = {"scripts": {".py"}, "references": {".md", ".json"}, "tests": {".py"}, "examples": {".json"}, "assets": {".docx"}}


def collect():
    result = [ROOT / name for name in ROOT_FILES if (ROOT / name).is_file()]
    for directory, suffixes in TREES.items():
        for path in sorted((ROOT / directory).rglob("*")):
            if path.is_file() and path.suffix in suffixes and "__pycache__" not in path.parts:
                result.append(path)
    return result


def package(output_dir):
    skill = (ROOT / "SKILL.md").read_text(encoding="utf-8")
    if not skill.startswith("---\n") or not re.search(r"^name: BNU-psy-thesis$", skill, re.M):
        raise ValueError("Invalid BNU-psy-thesis SKILL.md frontmatter")
    if not re.search(r"^description:", skill, re.M):
        raise ValueError("Missing skill description")
    files = collect()
    for path in files:
        if path.suffix == ".docx":
            continue
        text = path.read_text(encoding="utf-8")
        # Detect machine-specific absolute user paths; do not echo any secret contents.
        if re.search(r"(?i)[A-Z]:[\\/]Users[\\/](?!<|\{)[^\s\"']+", text):
            raise ValueError("Machine-specific user path in " + path.relative_to(ROOT).as_posix())
    target = output_dir / "BNU-psy-thesis.skill"
    manifest = output_dir / "manifest.json"
    output_dir.mkdir(parents=True, exist_ok=True)
    if target.exists() or manifest.exists() or (output_dir.resolve() == (ROOT / "dist").resolve() and any(output_dir.iterdir())):
        history = ROOT / "_local" / "release-history" / uuid.uuid4().hex
        history.mkdir(parents=True)
        for path in (target, manifest):
            if path.exists():
                shutil.move(str(path), str(history / path.name))
        if output_dir.resolve() == (ROOT / "dist").resolve():
            for path in list(output_dir.iterdir()):
                if path.is_dir():
                    shutil.move(str(path), str(history / path.name))
    records = []
    with target.open("xb") as handle:
        with zipfile.ZipFile(handle, "w", zipfile.ZIP_DEFLATED) as z:
            for file in files:
                relative = file.relative_to(ROOT).as_posix()
                data = file.read_bytes()
                z.writestr("BNU-psy-thesis/" + relative, data)
                records.append({"path": relative, "sha256": hashlib.sha256(data).hexdigest(), "bytes": len(data)})
    result = {"version": "2.4.3", "archive": target.name,
              "sha256": hashlib.sha256(target.read_bytes()).hexdigest(),
              "files": records, "excluded": ["_local", "dist", "private source documents", "credentials"]}
    with manifest.open("x", encoding="utf-8") as handle: json.dump(result, handle, ensure_ascii=False, indent=2)
    with zipfile.ZipFile(target) as z:
        if z.testzip() is not None: raise ValueError("Release ZIP failed integrity check")
        if any("/_local/" in n or "/dist/" in n for n in z.namelist()):
            raise ValueError("Private artifacts entered the release")
    return target, result


def main():
    p = argparse.ArgumentParser(description=__doc__)
    p.add_argument("--out", type=Path, required=True)
    args = p.parse_args()
    try: path, result = package(args.out)
    except (ValueError, OSError) as exc: p.exit(1, "ERROR: " + str(exc) + "\n")
    print("Packaged {} public files: {}".format(len(result["files"]), path))


if __name__ == "__main__": main()

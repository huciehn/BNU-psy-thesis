"""Restore and verify the bundled public source documents."""
import argparse
import base64
import hashlib
import json
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]


def digest(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()


def verify_sources(root):
    root = Path(root)
    manifest = json.loads((root / "references/sources.json").read_text(encoding="utf-8"))
    for item in manifest["sources"]:
        relative = item.get("path", "assets/sources/" + item["filename"])
        path = root / relative
        if not path.is_file():
            raise ValueError("Bundled source is missing: " + relative)
        if digest(path) != item["sha256"]:
            raise ValueError("Bundled source hash differs: " + relative)
    return len(manifest["sources"])


def ensure_sources(root=None):
    root = ROOT if root is None else Path(root)
    fallback_path = root / "references/source-fallback.json"
    if fallback_path.is_file():
        fallback = json.loads(fallback_path.read_text(encoding="utf-8"))
        for item in fallback["sources"]:
            target = root / "assets/sources" / item["filename"]
            if target.exists():
                if digest(target) != item["sha256"]:
                    raise ValueError("Refusing to replace a different source: " + target.name)
                continue
            if item.get("encoding") != "base64":
                raise ValueError("Unsupported source archive encoding")
            data = base64.b64decode(item["data"], validate=True)
            if hashlib.sha256(data).hexdigest() != item["sha256"]:
                raise ValueError("Source archive hash differs: " + item["filename"])
            target.parent.mkdir(parents=True, exist_ok=True)
            with target.open("xb") as handle:
                handle.write(data)
    return verify_sources(root)


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.parse_args()
    try:
        count = ensure_sources()
    except (OSError, ValueError) as exc:
        parser.exit(1, "ERROR: " + str(exc) + "\n")
    print("Verified {} bundled public sources".format(count))


if __name__ == "__main__":
    main()

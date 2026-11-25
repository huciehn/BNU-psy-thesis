"""Import private source material for local comparison; never change originals."""
import argparse
import hashlib
import json
import shutil
from pathlib import Path


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("source", type=Path)
    parser.add_argument("output", type=Path)
    args = parser.parse_args()
    args.output.mkdir(parents=True, exist_ok=True)
    manifest = []
    for src in sorted(args.source.rglob("*")):
        if not src.is_file() or src.name.startswith("~$"):
            continue
        digest = hashlib.sha256(src.read_bytes()).hexdigest()
        dst = args.output / src.name
        if dst.exists() and hashlib.sha256(dst.read_bytes()).hexdigest() != digest:
            raise SystemExit("Refusing to overwrite a different source: " + str(dst))
        if not dst.exists():
            shutil.copy2(src, dst)
        manifest.append({"filename": src.name, "sha256": digest})
        lines = ["# " + src.name]
        if src.suffix.lower() == ".docx":
            from docx import Document
            doc = Document(src)
            for i, p in enumerate(doc.paragraphs):
                if p.text.strip():
                    lines.append("[P{}] {}".format(i, p.text))
            for ti, table in enumerate(doc.tables):
                for ri, row in enumerate(table.rows):
                    lines.append("[T{}R{}] {}".format(
                        ti, ri, " | ".join(c.text for c in row.cells)))
        elif src.suffix.lower() == ".pdf":
            import fitz
            with fitz.open(src) as doc:
                for i, page in enumerate(doc):
                    lines.append("[PAGE {}]\n{}".format(i + 1, page.get_text()))
        else:
            continue
        text_path = args.output / (src.stem + ".md")
        if not text_path.exists():
            text_path.write_text("\n".join(lines), encoding="utf-8")
    target = args.output / "manifest.json"
    if target.exists():
        raise SystemExit("Manifest already exists; use a fresh output directory")
    target.write_text(json.dumps(manifest, ensure_ascii=False, indent=2), encoding="utf-8")
    print("Imported {} sources".format(len(manifest)))


if __name__ == "__main__":
    main()

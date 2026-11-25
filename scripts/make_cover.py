"""Build a detector-compatible cover from a user-supplied official DOCX sample."""
import argparse
from datetime import datetime
import json
from pathlib import Path
import re
import sys

from docx import Document

ROOT = Path(__file__).resolve().parents[1]


def make_blocks(args, logo_name):
    return [
        {"type": "image", "path": logo_name, "width_cm": 8.04},
        {"role": "cover_main_title", "text": "本科生毕业论文(设计)"},
        {"role": "blank", "text": ""},
        {"role": "blank", "text": ""},
        {"role": "cover_label", "text": "毕业论文（设计）题目："},
        {"role": "cover_title", "text": args.title},
        {"role": "blank", "text": ""},
        {"role": "blank", "text": ""},
        {"role": "blank", "text": ""},
        {"role": "cover_meta", "text": "部   院  系：" + args.department},
        {"role": "cover_meta", "text": "专       业：" + args.major},
        {"role": "cover_meta", "text": "学       号：" + args.student_id},
        {"role": "cover_meta", "text": "学 生 姓 名：" + args.student_name},
        {"role": "cover_meta", "text": "指 导 教 师：" + args.advisor},
        {"role": "cover_meta", "text": "指导教师职称：" + args.advisor_title},
        {"role": "cover_meta", "text": "指导教师单位：" + args.advisor_unit},
        {"role": "blank", "text": ""},
        {"role": "cover_date", "text": args.date},
    ]


def extract_logo(template, destination):
    document = Document(template)
    if not document.inline_shapes:
        raise ValueError("The official sample has no inline cover logo")
    shape = document.inline_shapes[0]
    relationship = shape._inline.graphic.graphicData.pic.blipFill.blip.embed
    destination.write_bytes(document.part.related_parts[relationship].blob)


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("template", type=Path, nargs="?", default=ROOT / "assets" / "cover-template.docx")
    parser.add_argument("--content", required=True, type=Path)
    parser.add_argument("--output", required=True, type=Path)
    parser.add_argument("--title", required=True)
    parser.add_argument("--student-id", required=True)
    parser.add_argument("--student-name", required=True)
    parser.add_argument("--advisor", required=True)
    parser.add_argument("--department", default="心理学部")
    parser.add_argument("--major", default="心理学")
    parser.add_argument("--advisor-title", default="")
    parser.add_argument("--advisor-unit", default="北京师范大学心理学部")
    parser.add_argument("--date", default=datetime.now().strftime("%Y年%m月%d日"))
    parser.add_argument("--replace-cover", action="store_true")
    args = parser.parse_args()
    try:
        if not re.fullmatch(r"\d{12}", args.student_id):
            raise ValueError("Student ID must contain exactly 12 digits")
        if args.output.exists():
            raise ValueError("Refusing to overwrite: " + str(args.output))
        content = json.loads(args.content.read_text(encoding="utf-8"))
        blocks = content["blocks"]
        if args.replace_cover:
            first_abstract = next((i for i, block in enumerate(blocks) if block.get("role") == "abstract_title_cn"), None)
            if first_abstract is None:
                raise ValueError("Cannot identify the existing cover boundary")
            blocks = blocks[first_abstract:]
        else:
            if any(block.get("role") in ("cover_title", "cover_main_title") for block in blocks):
                raise ValueError("Input already has a cover; use --replace-cover after reviewing it")
        for block in blocks:
            if block.get("type") == "image" and not Path(block["path"]).is_absolute():
                block["path"] = str((args.content.parent / block["path"]).resolve())
        logo = args.output.parent / "cover-logo.png"
        if logo.exists():
            raise ValueError("Refusing to overwrite: " + str(logo))
        args.output.parent.mkdir(parents=True, exist_ok=True)
        extract_logo(args.template, logo)
        content["blocks"] = make_blocks(args, logo.name) + blocks
        with args.output.open("x", encoding="utf-8") as handle:
            json.dump(content, handle, ensure_ascii=False, indent=2)
        print("Cover content: %s\nLogo: %s" % (args.output, logo))
    except Exception as exc:
        print("ERROR: " + str(exc), file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())

"""Build a plain DOCX and a role plan from supplied structured content."""
import argparse
import json
from pathlib import Path
from docx import Document
from docx.oxml import OxmlElement
from docx.shared import Cm
from lxml import etree
from docx_common import load_profile, parse_xml, paragraphs, read_docx, sha256_file, qn, NS


def build(content_path, output, plan_path):
    for path in (output, plan_path):
        if path.exists():
            raise ValueError("Refusing to overwrite: " + str(path))
    if len({content_path.resolve(), output.resolve(), plan_path.resolve()}) != 3:
        raise ValueError("Input and outputs must have different paths")
    content = json.loads(content_path.read_text(encoding="utf-8-sig"))
    profile = load_profile()
    d = Document()
    d.core_properties.author = ""
    d.core_properties.last_modified_by = ""
    role_nodes = []
    table_plan = {}
    for block in content["blocks"]:
        kind = block.get("type", "paragraph")
        if kind == "paragraph":
            role = block["role"]
            if role not in profile["styles"]:
                raise ValueError("Unknown role: " + role)
            p = d.add_paragraph(block["text"])
            item = {"role": role}
            if "spans" in block: item["spans"] = block["spans"]
            role_nodes.append((p._p, item))
        elif kind == "toc":
            p = d.add_paragraph(); r = p.add_run()
            for typ in ("begin", "separate", "end"):
                fc = OxmlElement("w:fldChar"); fc.set(qn("w:fldCharType"), typ)
                r._r.append(fc)
                if typ == "begin":
                    ins = OxmlElement("w:instrText"); ins.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
                    ins.text = ' TOC \\o "1-3" \\h \\z \\u '; r._r.append(ins)
            role_nodes.append((p._p, {"role": "toc1"}))
        elif kind == "table":
            rows = block["rows"]
            if not rows or not rows[0] or any(len(row) != len(rows[0]) for row in rows):
                raise ValueError("Tables must be nonempty rectangular row arrays")
            table = d.add_table(rows=len(rows), cols=len(rows[0]))
            tid = "t{:05d}".format(len(table_plan))
            table_plan[tid] = {"kind": block["kind"], "header_rows": block.get("header_rows", 1)}
            for row, values in zip(table.rows, rows):
                for cell, text in zip(row.cells, values):
                    cell.text = str(text)
                    role_nodes.append((cell.paragraphs[0]._p, {"role": "table_cell"}))
        elif kind == "image":
            image = (content_path.parent / block["path"]).resolve()
            if not image.is_file(): raise ValueError("Image file missing: " + str(image))
            p = d.add_paragraph()
            p.add_run().add_picture(str(image), width=Cm(block.get("width_cm", 10)))
            role_nodes.append((p._p, {"role": "image"}))
        elif kind == "formula":
            node = parse_xml(block["omml"].encode("utf-8"))
            if node.tag not in (qn("m:oMath"), qn("m:oMathPara")):
                raise ValueError("Formula must be valid OMML")
            p = d.add_paragraph(); p._p.append(node)
            role_nodes.append((p._p, {"role": "formula"}))
        else:
            raise ValueError("Unknown content block: " + kind)
    # python-docx's element subclass has a narrower xpath() signature than
    # ordinary lxml elements; direct iteration also retains object identity.
    all_ps = list(d._element.body.iter(qn("w:p")))
    plan = {"profile": profile["id"], "paragraphs": {}, "tables": table_plan}
    for node, item in role_nodes:
        plan["paragraphs"]["p{:05d}".format(all_ps.index(node))] = item
    output.parent.mkdir(parents=True, exist_ok=True)
    plan_path.parent.mkdir(parents=True, exist_ok=True)
    # Serialize to memory first; exclusive creation prevents accidental replacement.
    import io
    stream = io.BytesIO(); d.save(stream)
    with output.open("xb") as handle: handle.write(stream.getvalue())
    plan["input_sha256"] = sha256_file(output)
    with plan_path.open("x", encoding="utf-8") as handle:
        json.dump(plan, handle, ensure_ascii=False, indent=2)


def main():
    p = argparse.ArgumentParser(description=__doc__)
    p.add_argument("content", type=Path)
    p.add_argument("--output", type=Path, required=True)
    p.add_argument("--plan", type=Path, required=True)
    args = p.parse_args()
    try: build(args.content, args.output, args.plan)
    except (ValueError, KeyError, OSError) as exc: p.exit(1, "ERROR: " + str(exc) + "\n")
    print("Created a source DOCX and plan; run formatting and validation next")


if __name__ == "__main__": main()

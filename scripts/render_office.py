"""Render/update a new copy with explicitly selected local office software."""
import argparse
import hashlib
import json
import shutil
import subprocess
from pathlib import Path


def sha(path):
    return hashlib.sha256(Path(path).read_bytes()).hexdigest()


def render_once(source, output_dir, backend, executable=None):
    if output_dir.exists() and any(output_dir.iterdir()):
        raise ValueError("Use a new or empty output directory; no outputs are overwritten")
    output_dir.mkdir(parents=True, exist_ok=True)
    source = source.resolve()
    before = sha(source)
    updated = output_dir.resolve() / "updated.docx"
    pdf = output_dir.resolve() / "rendered.pdf"
    pages = None
    section_pages = []
    if backend in ("word", "wps"):
        try:
            import win32com.client
        except ImportError as exc:
            raise ValueError("Windows rendering requires pywin32 and installed Word/WPS") from exc
        app = win32com.client.DispatchEx("Word.Application" if backend == "word" else "KWPS.Application")
        owned = False
        doc = None
        try:
            if app.Documents.Count:
                raise ValueError("Office returned a session with existing documents; refusing to operate on that session")
            owned = True
            app.Visible = False
            app.DisplayAlerts = 0
            doc = app.Documents.Open(str(source), ReadOnly=True, AddToRecentFiles=False)
            doc.Fields.Update()
            for i in range(1, doc.TablesOfContents.Count + 1):
                doc.TablesOfContents.Item(i).TabLeader = 5
                doc.TablesOfContents.Item(i).Update()
            doc.Repaginate()
            # F9 can change pagination: refresh cached page numbers after repagination.
            for i in range(1, doc.TablesOfContents.Count + 1):
                doc.TablesOfContents.Item(i).UpdatePageNumbers()
            for i in range(1, doc.Sections.Count + 1):
                sec = doc.Sections.Item(i)
                rng = sec.Range.Duplicate
                rng.Collapse(1)
                section_pages.append({"section": i, "physical_page": int(rng.Information(3)),
                                      "numbered_page": int(rng.Information(1)),
                                      "section_start": int(sec.PageSetup.SectionStart)})
            try:
                doc.SaveAs2(str(updated), FileFormat=12)
            except AttributeError:
                doc.SaveAs(str(updated), FileFormat=12)
            doc.ExportAsFixedFormat(str(pdf), 17)
            pages = int(doc.ComputeStatistics(2))
        finally:
            if doc is not None: doc.Close(SaveChanges=False)
            if owned: app.Quit()
    else:
        if not executable:
            raise ValueError("LibreOffice backend requires --executable with an explicit executable path")
        # Isolated application profile stays inside this task directory.
        profile = output_dir.resolve() / "lo-profile"
        copy = output_dir.resolve() / "source.docx"
        shutil.copy2(source, copy)
        command = [str(executable), "-env:UserInstallation=" + profile.as_uri(),
                   "--headless", "--convert-to", "pdf", "--outdir", str(output_dir.resolve()), str(copy)]
        result = subprocess.run(command, capture_output=True, text=True, timeout=180)
        (output_dir / "office.log").write_text(result.stdout + "\n" + result.stderr, encoding="utf-8")
        generated = output_dir / "source.pdf"
        if result.returncode or not generated.is_file():
            raise ValueError("LibreOffice failed; see office.log")
        # A conversion is a render, not proof that the editable file's TOC was saved.
        shutil.copy2(generated, pdf)
        updated = None
    if sha(source) != before:
        raise ValueError("Original file changed unexpectedly")
    if not pdf.is_file() or pdf.stat().st_size == 0:
        raise ValueError("Office did not produce a nonempty PDF")
    report = {"backend": backend, "source_sha256": before, "rendered_pdf_sha256": sha(pdf),
              "updated_docx_sha256": sha(updated) if updated else None, "page_count": pages,
              "sections": section_pages,
              "editable_fields_saved": updated is not None, "all_pages_visual": False,
              "status": "rendered_pending_visual_and_structural_review"}
    (output_dir / "render.json").write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    return report


def pad_before_body(source, output, plan):
    """Insert a real blank leaf before body numbering restarts (no prose edits)."""
    import zipfile
    from lxml import etree
    from docx_common import read_docx, NS, qn, xml_bytes
    parts = read_docx(source)
    root = etree.fromstring(parts["word/document.xml"])
    body = root.find("w:body", NS)
    pid = next((k for k, v in plan["paragraphs"].items() if v["role"] == "heading1"), None)
    if not pid: raise ValueError("Cannot identify the first body heading from the reviewed plan")
    mark = root.find('.//w:bookmarkStart[@w:name="BNU-psy-thesis_' + pid + '"]', NS)
    if mark is None: raise ValueError("Body bookmark is missing")
    paragraph = next(mark.iterancestors(qn("w:p")), None)
    if paragraph is None or paragraph.getparent() is not body:
        raise ValueError("The body must start in a top-level paragraph")
    previous = paragraph.getprevious()
    section = previous.find("w:pPr/w:sectPr", NS) if previous is not None else None
    if section is None: raise ValueError("Cannot safely locate the front/body section boundary")
    section.getparent().remove(section)
    first = etree.Element(qn("w:p"))
    pr = etree.SubElement(first, qn("w:pPr")); etree.SubElement(pr, qn("w:pageBreakBefore"))
    second = etree.Element(qn("w:p")); etree.SubElement(second, qn("w:pPr")).append(section)
    index = body.index(paragraph)
    body.insert(index, first); body.insert(index + 1, second)
    parts["word/document.xml"] = xml_bytes(root)
    with output.open("xb") as handle:
        with zipfile.ZipFile(handle, "w", zipfile.ZIP_DEFLATED) as z:
            for name, data in parts.items(): z.writestr(name, data)


def render(source, output_dir, backend, executable=None, plan_path=None):
    if not plan_path:
        return render_once(source, output_dir, backend, executable)
    if backend == "libreoffice":
        raise ValueError("Automatic physical-page repair currently needs Word/WPS; use PDF review with LibreOffice")
    if output_dir.exists() and any(output_dir.iterdir()):
        raise ValueError("Use a new or empty output directory")
    output_dir.mkdir(parents=True, exist_ok=True)
    plan = json.loads(plan_path.read_text(encoding="utf-8-sig"))
    first = render_once(source, output_dir / "pass1", backend, executable)
    body = next((s for s in first["sections"] if s["numbered_page"] == 1 and s["section"] > 1), None)
    result, final_dir = first, output_dir / "pass1"
    if body and body["physical_page"] % 2 == 0:
        padded = output_dir / "physical-padding.docx"
        pad_before_body(source, padded, plan)
        result = render_once(padded, output_dir / "pass2", backend, executable)
        final_dir = output_dir / "pass2"
    result["physical_odd_sections_verified"] = all(s["physical_page"] % 2 == 1 for s in result["sections"] if s["section_start"] == 4)
    if not result["physical_odd_sections_verified"]:
        raise ValueError("Physical odd-page requirement still fails; inspect pass reports before further editing")
    result["final_directory"] = final_dir.name
    (output_dir / "render-summary.json").write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")
    return result


def main():
    p = argparse.ArgumentParser(description=__doc__)
    p.add_argument("source", type=Path)
    p.add_argument("--output-dir", required=True, type=Path)
    p.add_argument("--backend", required=True, choices=("word", "wps", "libreoffice"))
    p.add_argument("--executable", type=Path)
    p.add_argument("--plan", type=Path, help="Use the reviewed plan to verify and repair physical odd-page starts")
    args = p.parse_args()
    try: result = render(args.source, args.output_dir, args.backend, args.executable, args.plan)
    except Exception as exc: p.exit(1, "ERROR: " + str(exc) + "\n")
    print(json.dumps(result, indent=2))


if __name__ == "__main__": main()

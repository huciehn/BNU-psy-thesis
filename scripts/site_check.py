"""Configure the university detector, test login, or upload an authorized DOCX."""
import argparse
import json
from pathlib import Path
import re
import sys
import time

import requests

ROOT = Path(__file__).resolve().parents[1]
CREDENTIALS = Path.home() / ".claude" / "secrets" / "BNU-psy-thesis-site.json"


def load_public_config():
    return json.loads((ROOT / "references/site.json").read_text(encoding="utf-8"))


def validate_credentials(student_id, email):
    if not re.fullmatch(r"\d{12}", student_id):
        raise ValueError("Student ID must contain exactly 12 digits")
    if not re.fullmatch(r"[^\s@]+@[^\s@]+\.[^\s@]+", email):
        raise ValueError("Email address is invalid")


def login(session, entry, credentials):
    response = session.post(
        entry.rstrip("/") + "/Home/Login",
        data={"txtStudentId": credentials["student_id"], "txtEmail": credentials["email"]},
        timeout=30,
    )
    response.raise_for_status()
    if "退出" not in response.text:
        raise ValueError("Detector login was not accepted")
    return response


def save_credentials(config):
    CREDENTIALS.parent.mkdir(parents=True, exist_ok=True)
    with CREDENTIALS.open("w", encoding="utf-8") as handle:
        json.dump(config, handle, ensure_ascii=False, indent=2)
    try:
        CREDENTIALS.chmod(0o600)
    except OSError:
        pass


def reports(session, entry):
    response = session.get(entry.rstrip("/") + "/Home/GetReportList", timeout=30)
    response.raise_for_status()
    return response.json()


def upload_and_download(session, entry, document, output_dir, language, credentials):
    if document.suffix.lower() != ".docx":
        raise ValueError("The detector requires a standard DOCX file")
    if document.stat().st_size > 10 * 1024 * 1024:
        raise ValueError("The detector accepts files up to 10 MiB")
    before = {item["fileName"] for item in reports(session, entry)}
    with document.open("rb") as stream:
        response = session.post(
            entry,
            files={"paperFile": (document.name, stream, "application/vnd.openxmlformats-officedocument.wordprocessingml.document")},
            data={
                "language": language,
                "email": credentials["email"],
                "studentId": credentials["student_id"],
                "paperTitle": document.stem,
                "paperType": "毕业论文",
                "additionalInfo": "",
            },
            timeout=180,
        )
    response.raise_for_status()
    result = response.json()
    if not result.get("success"):
        raise ValueError(result.get("message", "Upload was rejected"))
    deadline = time.monotonic() + 300
    new_items = []
    while time.monotonic() < deadline:
        new_items = [item for item in reports(session, entry) if item["fileName"] not in before]
        if any(item["fileName"].endswith(".txt") for item in new_items):
            break
        time.sleep(5)
    if not new_items:
        raise TimeoutError("No report appeared within 300 seconds")
    output_dir.mkdir(parents=True, exist_ok=True)
    saved = []
    for item in new_items:
        target = output_dir / item["fileName"]
        if target.exists():
            raise ValueError("Refusing to overwrite report: " + str(target))
        response = session.get(item["downloadUrl"], timeout=60)
        response.raise_for_status()
        target.write_bytes(response.content)
        saved.append(str(target))
    return saved


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--setup", action="store_true")
    parser.add_argument("--check", action="store_true")
    parser.add_argument("--upload", type=Path)
    parser.add_argument("--report-dir", type=Path)
    parser.add_argument("--language", choices=("中文", "英文"), default="中文")
    parser.add_argument("--yes", action="store_true", help="Confirm that this DOCX may be sent to the configured detector")
    args = parser.parse_args()
    public = load_public_config()
    try:
        if args.setup:
            entry = input("检测入口 [" + public["entry"] + "]: ").strip() or public["entry"]
            student_id = input("12位学号: ").strip()
            email = input("邮箱: ").strip()
            validate_credentials(student_id, email)
            credentials = {"student_id": student_id, "email": email, "entry": entry}
            session = requests.Session()
            login(session, entry, credentials)
            save_credentials(credentials)
            print(json.dumps({"status": "configured", "entry": entry}, ensure_ascii=False))
            return 0
        credentials = json.loads(CREDENTIALS.read_text(encoding="utf-8"))
        validate_credentials(credentials["student_id"], credentials["email"])
        entry = credentials.get("entry", public["entry"])
        session = requests.Session()
        login(session, entry, credentials)
        if args.check:
            print(json.dumps({"status": "authenticated", "entry": entry}, ensure_ascii=False))
            return 0
        if args.upload:
            if not args.yes:
                raise ValueError("Upload requires --yes after the user authorizes this document transfer")
            if not args.report_dir:
                raise ValueError("--report-dir is required")
            saved = upload_and_download(session, entry, args.upload.resolve(), args.report_dir.resolve(), args.language, credentials)
            print(json.dumps({"status": "report_downloaded", "files": saved}, ensure_ascii=False))
            return 0
        parser.error("Choose --setup, --check, or --upload")
    except Exception as exc:
        print(json.dumps({"status": "error", "message": str(exc)}, ensure_ascii=False), file=sys.stderr)
        return 1


if __name__ == "__main__":
    sys.exit(main())

#!/usr/bin/env python3
"""Local web form server for the pension packet generator (stdlib only).

Serves pension_form.html and turns a JSON POST into the same three DOCX
outputs as pension_packet_generator.py, returned as one ZIP download.
No new dependencies: uses http.server + pension_packet_generator functions.
"""
from __future__ import annotations

import io
import json
import shutil
import tempfile
import webbrowser
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import urlparse
from zipfile import ZipFile

import pension_packet_generator as gen

HERE = Path(__file__).resolve().parent
FORM = HERE / "pension_form.html"
PORT = 8000


def parse_date_or_none(value: str, label: str, required: bool = False):
    value = (value or "").strip()
    if not value:
        if required:
            raise ValueError(f"{label} is required.")
        return None
    # HTML date inputs give YYYY-MM-DD; also accept the CLI formats.
    return gen.parse_date(value, label, required=required)


def build_payload(body: dict):
    data = {
        "name": (body.get("name") or "").strip(),
        "designation": (body.get("designation") or "").strip(),
        "employee_code": (body.get("employee_code") or "").strip(),
        "gpf": (body.get("gpf") or "").strip(),
        "dob": parse_date_or_none(body.get("dob", ""), "Date of Birth", required=True),
        "joining": parse_date_or_none(body.get("joining", ""), "Date of Joining", required=True),
        "retirement": parse_date_or_none(body.get("retirement", ""), "Date of Superannuation", required=True),
        "issue_date": parse_date_or_none(body.get("issue_date", ""), "Date of Issue", required=True),
        "headmaster": (body.get("headmaster") or "").strip(),
        "pan": gen.normalize_pan(body.get("pan") or ""),
        "mobile": gen.normalize_mobile(body.get("mobile") or ""),
        "email": (body.get("email") or "").strip(),
        "height": (body.get("height") or "").strip(),
        "blood_group": (body.get("blood_group") or "").strip(),
        "address": (body.get("address") or "").strip(),
        "identification_mark": (body.get("identification_mark") or "").strip(),
        "treasury": (body.get("treasury") or "").strip(),
    }
    if not data["name"]:
        raise ValueError("Employee full name is required.")
    if not data["designation"]:
        raise ValueError("Designation is required.")
    if data["retirement"] < data["joining"]:
        raise ValueError("Retirement date cannot be earlier than the joining date.")

    calculated = gen.qualifying_service(data["joining"], data["retirement"])
    data["qualifying_service"] = (body.get("qualifying_service") or "").strip() or calculated

    periods = []
    for p in body.get("periods") or []:
        f = parse_date_or_none(p.get("from", ""), "Service period From date", required=True)
        t = parse_date_or_none(p.get("to", ""), "Service period To date", required=True)
        page = (p.get("page") or "").strip()
        if not page:
            raise ValueError("Each service period needs a Service Book page number.")
        if t < f:
            raise ValueError("A service period ends before it starts.")
        if f < data["joining"] or t > data["retirement"]:
            raise ValueError("A service period falls outside the joining–retirement dates.")
        periods.append({"from": f, "to": t, "page": page})
    ordered = sorted(periods, key=lambda p: (p["from"], p["to"]))
    for prev, cur in zip(ordered, ordered[1:]):
        if prev["to"] >= cur["from"]:
            raise ValueError("Two service periods overlap.")

    family = []
    for m in body.get("family") or []:
        name = (m.get("name") or "").strip()
        if not name:
            continue  # skip blank rows rather than failing
        dob_raw = (m.get("dob") or "").strip()
        dob_fmt = gen.fmt_date(parse_date_or_none(dob_raw, "Family member DOB")) if dob_raw else ""
        family.append({
            "name": name, "dob": dob_fmt,
            "marital_status": (m.get("marital_status") or "").strip(),
            "relation": (m.get("relation") or "").strip(),
            "disability": (m.get("disability") or "No").strip() or "No",
            "blood_group": (m.get("blood_group") or "").strip(),
            "height": (m.get("height") or "").strip(),
            "id_proof": (m.get("id_proof") or "").strip(),
        })

    enclosed = set()
    for n in body.get("enclosed") or []:
        try:
            num = int(n)
        except (TypeError, ValueError):
            continue
        if 1 <= num <= len(gen.ENCLOSURES):
            enclosed.add(num)
    return data, periods, family, enclosed


class Handler(BaseHTTPRequestHandler):
    server_version = "PensionForm/1.0"

    def _send(self, code: int, content: bytes, ctype: str, filename: str | None = None):
        self.send_response(code)
        self.send_header("Content-Type", ctype)
        self.send_header("Content-Length", str(len(content)))
        if filename:
            self.send_header("Content-Disposition", f'attachment; filename="{filename}"')
        self.end_headers()
        self.wfile.write(content)

    def do_GET(self):
        path = urlparse(self.path).path
        if path in ("/", "/pension_form.html"):
            if not FORM.exists():
                self._send(500, b"pension_form.html not found.", "text/plain")
                return
            self._send(200, FORM.read_bytes(), "text/html; charset=utf-8")
        elif path == "/api/enclosures":
            payload = json.dumps([
                {"n": i + 1, "label": label} for i, label in enumerate(gen.ENCLOSURES)
            ]).encode()
            self._send(200, payload, "application/json")
        else:
            self._send(404, b"Not found.", "text/plain")

    def do_POST(self):
        if urlparse(self.path).path != "/api/generate":
            self._send(404, b"Not found.", "text/plain")
            return
        try:
            length = int(self.headers.get("Content-Length") or 0)
            body = json.loads(self.rfile.read(length).decode("utf-8") or "{}")
            data, periods, family, enclosed = build_payload(body)
        except Exception as exc:
            self._send(400, json.dumps({"error": str(exc)}).encode(), "application/json")
            return
        if not gen.RETIREMENT_TEMPLATE.exists() or not gen.KRITAGYATA_TEMPLATE.exists():
            self._send(500, json.dumps({"error": "DOCX templates not found beside the script."}).encode(),
                       "application/json")
            return
        try:
            tmp = Path(tempfile.mkdtemp(prefix="pension_"))
            safe = gen.slugify(data["name"])
            r_out = tmp / f"Retirement_Certificates_{safe}.docx"
            k_out = tmp / f"Kritagyata_Form_{safe}.docx"
            c_out = tmp / f"Complete_Pension_Packet_{safe}.docx"
            gen.fill_retirement_template(data, periods, r_out)
            gen.fill_kritagyata_template(data, family, enclosed, k_out)
            gen.combine_documents(r_out, k_out, c_out)
            # Persist a copy in Generated_Pension_Packets (same as CLI).
            out_dir = HERE / "Generated_Pension_Packets"
            out_dir.mkdir(exist_ok=True)
            for p in (r_out, k_out, c_out):
                shutil.copy2(p, out_dir / p.name)
            buf = io.BytesIO()
            with ZipFile(buf, "w") as z:
                for p in (r_out, k_out, c_out):
                    z.write(p, p.name)
            shutil.rmtree(tmp, ignore_errors=True)
            self._send(200, buf.getvalue(), "application/zip", f"Pension_Packet_{safe}.zip")
        except Exception as exc:
            self._send(500, json.dumps({"error": str(exc)}).encode(), "application/json")

    def log_message(self, *args):
        pass  # keep console quiet


def main() -> int:
    server = ThreadingHTTPServer(("127.0.0.1", PORT), Handler)
    url = f"http://127.0.0.1:{PORT}/"
    print(f"Pension form running at {url}")
    print("Fill the form in your browser. Press Ctrl+C to stop.")
    try:
        webbrowser.open(url)
    except Exception:
        pass
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\nStopped.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

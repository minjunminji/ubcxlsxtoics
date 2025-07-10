#!/usr/bin/env python3
"""
Serverless Python function for Vercel.
Converts the uploaded UBC Workday “View My Courses” Excel export
into an .ics calendar file and streams it back to the caller.
"""

from __future__ import annotations
import sys, uuid, json, re
from datetime import datetime, timedelta
from io import BytesIO
from http.server import BaseHTTPRequestHandler
from urllib.parse import parse_qs
import cgi
import traceback

# Lazy heavy imports
try:
    import pandas as pd  # type: ignore
    from dateutil import parser as date_parser  # type: ignore
except Exception:
    pd = None  # type: ignore
    date_parser = None  # type: ignore

VTIMEZONE_BLOCK = """BEGIN:VTIMEZONE
TZID:America/Vancouver
BEGIN:STANDARD
TZOFFSETFROM:-0700
TZOFFSETTO:-0800
TZNAME:PST
DTSTART:19701101T020000
RRULE:FREQ=YEARLY;BYMONTH=11;BYDAY=1SU
END:STANDARD
BEGIN:DAYLIGHT
TZOFFSETFROM:-0800
TZOFFSETTO:-0700
TZNAME:PDT
DTSTART:19700308T020000
RRULE:FREQ=YEARLY;BYMONTH=3;BYDAY=2SU
END:DAYLIGHT
END:VTIMEZONE"""

LABOR_DAY_2025 = datetime(2025, 9, 1).date()

DAY_MAP = {
    "Mon": "MO", "Tue": "TU", "Wed": "WE", "Thu": "TH",
    "Fri": "FR", "Sat": "SA", "Sun": "SU",
}

def make_uid() -> str:
    return f"{uuid.uuid4()}@ubc-xlsx-to-ics"

def make_dtstamp() -> str:
    try:
        return datetime.now(datetime.UTC).strftime("%Y%m%dT%H%M%SZ")
    except AttributeError:  # <3.11 fallback
        return datetime.utcnow().strftime("%Y%m%dT%H%M%SZ")

def parse_meeting_pattern(pattern: str):
    # Same logic as earlier index.py
    try:
        if not pattern or pattern.strip() == "":
            return (None,) * 6
        parts = [p.strip() for p in pattern.split("|")]
        if len(parts) < 4:
            parts.extend([''] * (4 - len(parts)))
        date_range, days, times, location = parts[0], parts[1], parts[2], parts[3]
        date_match = re.match(r'(\d{4}-\d{2}-\d{2})\s*[-–]\s*(\d{4}-\d{2}-\d{2})', date_range)
        if not date_match:
            return (None,) * 6
        start_date, end_date = date_match.groups()
        if " - " in times:
            start_time, end_time = [t.strip() for t in times.split(" - ")]
        elif "-" in times:
            start_time, end_time = [t.strip() for t in times.split("-")]
        else:
            return (None,) * 6
        return start_date, end_date, days, start_time, end_time, location
    except Exception:
        return (None,) * 6

def convert_excel_to_ics(file_content: bytes):
    global pd, date_parser
    if pd is None or date_parser is None:
        try:
            import importlib
            pd = importlib.import_module("pandas")  # type: ignore
            date_parser = importlib.import_module("dateutil.parser")  # type: ignore
        except ModuleNotFoundError as ie:
            return None, f"Missing Python dependency: {ie.name}"

    try:
        buf = BytesIO(file_content)
        raw = pd.read_excel(buf, header=None, engine='openpyxl')
        mask = raw.apply(lambda r: r.astype(str).str.contains(r"Course Listing", case=False, na=False).any(), axis=1)
        hdr_idx = int(mask.idxmax()) if mask.any() else 2
        buf.seek(0)
        df = pd.read_excel(buf, header=hdr_idx, engine='openpyxl')

        events = []
        for _, row in df.iterrows():
            section_info = str(row.get("Section", "")).strip()
            patterns = str(row.get("Meeting Patterns", "")).strip()
            instr = str(row.get("Instructor", "")).strip()
            if not (section_info and patterns):
                continue
            summary = section_info.replace("_V", "")
            for line in patterns.split("\n"):
                sd, ed, days, st, et, loc = parse_meeting_pattern(line)
                if not all([sd, ed, days, st, et]):
                    continue
                byday = [DAY_MAP[d] for d in DAY_MAP if d in days]
                if not byday:
                    continue
                dt_start0 = date_parser.parse(f"{sd} {st}")
                dt_end0 = date_parser.parse(f"{sd} {et}")
                until_dt = date_parser.parse(f"{ed} {et}")
                exdate = ""
                if "MO" in byday and LABOR_DAY_2025 >= dt_start0.date() <= until_dt.date():
                    labor_dt = datetime.combine(LABOR_DAY_2025, dt_start0.time())
                    exdate = f"EXDATE;TZID=America/Vancouver:{labor_dt.strftime('%Y%m%dT%H%M%S')}\n"
                events.append(f"""BEGIN:VEVENT
UID:{make_uid()}
DTSTAMP:{make_dtstamp()}
SUMMARY:{summary}
DTSTART;TZID=America/Vancouver:{dt_start0.strftime('%Y%m%dT%H%M%S')}
DTEND;TZID=America/Vancouver:{dt_end0.strftime('%Y%m%dT%H%M%S')}
RRULE:FREQ=WEEKLY;BYDAY={','.join(byday)};UNTIL={until_dt.strftime('%Y%m%dT%H%M%S')}
{exdate}LOCATION:{loc}
DESCRIPTION:Instructor: {instr}\\nTime: {days} {st}-{et}
END:VEVENT""")
        if not events:
            return None, "No events found in the spreadsheet"
        cal = "\n".join(["BEGIN:VCALENDAR", "VERSION:2.0", "CALSCALE:GREGORIAN", VTIMEZONE_BLOCK, *events, "END:VCALENDAR"])
        return cal, None
    except Exception as e:
        return None, traceback.format_exc(limit=4)

class handler(BaseHTTPRequestHandler):
    def do_POST(self):
        try:
            ctype, pdict = cgi.parse_header(self.headers.get('content-type'))
            if ctype != 'multipart/form-data':
                self._err(400, 'Content-Type must be multipart/form-data')
                return
            pdict['boundary'] = bytes(pdict['boundary'], 'utf-8')
            fields = cgi.parse_multipart(self.rfile, pdict)
            fdata = fields.get('file')
            if not fdata:
                self._err(400, 'No file field in form')
                return
            ics, err = convert_excel_to_ics(fdata[0])
            if err:
                self._err(400, f"Conversion error: {err}")
                return
            self.send_response(200)
            self.send_header('Content-Type', 'text/calendar')
            self.send_header('Content-Disposition', 'attachment; filename="courses.ics"')
            self.end_headers()
            self.wfile.write(ics.encode('utf-8'))
        except Exception as e:
            tb = traceback.format_exc()
            print(tb, file=sys.stderr)
            self._err(500, 'Internal server error', tb)

    def _err(self, code, msg, tb=None):
        self.send_response(code)
        self.send_header('Content-Type', 'application/json')
        self.end_headers()
        payload = {'error': msg}
        if tb:
            payload['traceback'] = tb
        self.wfile.write(json.dumps(payload).encode()) 
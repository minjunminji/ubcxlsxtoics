#!/usr/bin/env python3
"""
Convert "View My Courses" .xlsx exports from Workday-UBC into
RFC 5545 .ics files for Google Calendar.
"""

from __future__ import annotations
import sys, uuid, re, json
from datetime import datetime, timedelta
from io import BytesIO

# We will import heavy libraries lazily inside functions to avoid import-time failures
try:
    import pandas as pd  # type: ignore
    from dateutil import parser as date_parser  # type: ignore
except Exception:
    # We will handle missing dependencies later during function call
    pd = None  # type: ignore
    date_parser = None  # type: ignore

from http.server import BaseHTTPRequestHandler
import json
from urllib.parse import parse_qs
import cgi
import traceback
# remove tempfile import as it's no longer needed
# import tempfile

# Keep your existing helper functions and constants
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

# Only exclude Labor Day
LABOR_DAY_2025 = datetime(2025, 9, 1).date()

DAY_MAP = {
    "Mon": "MO", "Tue": "TU", "Wed": "WE", "Thu": "TH",
    "Fri": "FR", "Sat": "SA", "Sun": "SU",
}

def make_uid() -> str:
    return f"{uuid.uuid4()}@ubc-xlsx-to-ics"

def make_dtstamp() -> str:
    """Generate a timestamp in UTC format per RFC 5545"""
    try:
        # Python 3.11+ preferred method
        return datetime.now(datetime.UTC).strftime("%Y%m%dT%H%M%SZ")
    except AttributeError:
        # Fallback for older Python versions
        return datetime.utcnow().strftime("%Y%m%dT%H%M%SZ")

def parse_meeting_pattern(pattern: str):
    """
    Split the Workday "Meeting Patterns" field:
        2026-01-06 – 2026-04-09 | Tue, Thu | 15:30PM-17:00PM | MacLeod 242
        or
        2025-09-05 - 2025-11-28 | Fri (Alternate weeks) | 4:00 p.m. - 6:00 p.m. | ESB-Floor 1-Room 1013
    →  (start_date, end_date, days, start_time, end_time, location)
    """
    import re
    try:
        # Check for empty pattern
        if not pattern or pattern.strip() == '':
            return (None,) * 6
            
        parts = [p.strip() for p in pattern.split("|")]
        # Handle missing location field
        if len(parts) < 4:
            parts.extend([''] * (4 - len(parts)))  # Pad with empty strings if needed
        date_range, days, times, location = parts[0], parts[1], parts[2], parts[3]
        
        # Try to extract start and end dates using a regex that matches the exact format
        # This handles both formats: "2025-09-05 - 2025-11-28" and "2026-01-06 – 2026-04-09"
        date_match = re.match(r'(\d{4}-\d{2}-\d{2})\s*[-–]\s*(\d{4}-\d{2}-\d{2})', date_range)
        if date_match:
            start_date, end_date = date_match.groups()
        else:
            print(f"DEBUG: Invalid date range format: '{date_range}'", file=sys.stderr)
            return (None,) * 6
        
        # Split time range - handle both formats: "4:00 p.m. - 6:00 p.m." and "15:30PM-17:00PM"
        if " - " in times:
            start_time, end_time = [t.strip() for t in times.split(" - ")]
        elif "-" in times:
            start_time, end_time = [t.strip() for t in times.split("-")]
        else:
            print(f"DEBUG: Invalid time format: '{times}'", file=sys.stderr)
            return (None,) * 6
            
        return start_date, end_date, days, start_time, end_time, location
    except Exception as e:
        print(f"DEBUG: Error parsing pattern '{pattern}': {e}", file=sys.stderr)
        return (None,) * 6

def convert_excel_to_ics(file_content: bytes):
    try:
        # Ensure pandas and dateutil are available
        global pd, date_parser
        if pd is None or date_parser is None:
            try:
                import importlib
                pd = importlib.import_module("pandas")  # type: ignore
                date_parser = importlib.import_module("dateutil.parser")  # type: ignore
            except ModuleNotFoundError as ie:
                return None, f"Missing Python dependency: {ie.name}. Please ensure it is in requirements.txt"

        # Process file in-memory instead of writing to disk
        file_buffer = BytesIO(file_content)
        
        raw = pd.read_excel(file_buffer, header=None, engine='openpyxl')
        mask = raw.apply(lambda r: r.astype(str).str.contains(r"Course Listing", case=False, na=False).any(), axis=1)
        hdr_idx = int(mask.idxmax()) if mask.any() else 2
        
        # Reset buffer's position to the beginning to be read again
        file_buffer.seek(0)
        df = pd.read_excel(file_buffer, header=hdr_idx, engine='openpyxl')
        
        events = []
        for _, row in df.iterrows():
            section_info = str(row.get("Section", "")).strip()
            patterns = str(row.get("Meeting Patterns", "")).strip()
            instr = str(row.get("Instructor", "")).strip()
            
            if not (section_info and patterns): continue
            
            summary = section_info.replace("_V", "")
            
            for line in patterns.split("\n"):
                sd, ed, days, st, et, loc = parse_meeting_pattern(line)
                if not all([sd, ed, days, st, et]): continue
                byday = [DAY_MAP[d] for d in DAY_MAP if d in days]
                if not byday: continue

                dt_start0 = date_parser.parse(f"{sd} {st}")
                dt_end0 = date_parser.parse(f"{sd} {et}")
                until_dt = date_parser.parse(f"{ed} {et}")
                
                exdate = ""
                if "MO" in byday:
                    start_date = dt_start0.date()
                    weekday_diff = (0 - start_date.weekday()) % 7
                    first_monday = start_date + timedelta(days=weekday_diff)
                    if first_monday == LABOR_DAY_2025:
                        dt_start0 += timedelta(days=7)
                        dt_end0 += timedelta(days=7)
                    if LABOR_DAY_2025 >= dt_start0.date() and LABOR_DAY_2025 <= until_dt.date():
                        labor_day_dt = datetime.combine(LABOR_DAY_2025, dt_start0.time())
                        labor_day_str = labor_day_dt.strftime("%Y%m%dT%H%M%S")
                        exdate = f"EXDATE;TZID=America/Vancouver:{labor_day_str}\n"
                
                dts = dt_start0.strftime("%Y%m%dT%H%M%S")
                dte = dt_end0.strftime("%Y%m%dT%H%M%S")
                until = until_dt.strftime("%Y%m%dT%H%M%S")
                events.append(f"""BEGIN:VEVENT
UID:{make_uid()}
DTSTAMP:{make_dtstamp()}
SUMMARY:{summary}
DTSTART;TZID=America/Vancouver:{dts}
DTEND;TZID=America/Vancouver:{dte}
RRULE:FREQ=WEEKLY;BYDAY={','.join(byday)};UNTIL={until}
{exdate}LOCATION:{loc}
DESCRIPTION:Instructor: {instr}\\nTime: {days} {st}-{et}
END:VEVENT""")
        
        if not events: return None, "No events found"
        
        cal = "\n".join(["BEGIN:VCALENDAR", "VERSION:2.0", "CALSCALE:GREGORIAN", VTIMEZONE_BLOCK, *events, "END:VCALENDAR"])
        return cal, None
    except Exception as e:
        import traceback, textwrap
        tb = traceback.format_exc()
        return None, textwrap.shorten(tb, width=1000)


# This is the Vercel Serverless Function handler
class handler(BaseHTTPRequestHandler):
    def do_POST(self):
        try:
            # 1. Parse the multipart form data
            content_type, pdict = cgi.parse_header(self.headers.get('content-type'))
            if content_type != 'multipart/form-data':
                self.send_error_response(400, 'Invalid content type: must be multipart/form-data')
                return

            pdict['boundary'] = bytes(pdict['boundary'], "utf-8")
            
            fields = cgi.parse_multipart(self.rfile, pdict)
            file_field = fields.get('file')
            
            if not file_field or not file_field[0]:
                self.send_error_response(400, 'File data not found in request')
                return
            
            file_content = file_field[0]
            
            # 2. Convert the file
            ics_data, error = convert_excel_to_ics(file_content)

            if error:
                # This error comes from a caught exception inside the conversion function
                self.send_error_response(400, f"Conversion Error: {error}")
                return

            # 3. Send the successful response
            self.send_response(200)
            self.send_header('Content-type', 'text/calendar')
            self.send_header('Content-Disposition', 'attachment; filename="courses.ics"')
            self.end_headers()
            self.wfile.write(ics_data.encode('utf-8'))

        except Exception as e:
            # This is the crucial part: catch ANY other exception
            tb_str = traceback.format_exc()
            print(f"--- UNHANDLED EXCEPTION --- \n{tb_str}", file=sys.stderr)
            self.send_error_response(500, "An internal server error occurred.", traceback=tb_str)

    def send_error_response(self, code, message, traceback=None):
        self.send_response(code)
        self.send_header('Content-type', 'application/json')
        self.end_headers()
        error_payload = {'error': message}
        if traceback:
            error_payload['traceback'] = traceback
        self.wfile.write(json.dumps(error_payload).encode('utf-8'))

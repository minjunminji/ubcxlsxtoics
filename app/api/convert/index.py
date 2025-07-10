#!/usr/bin/env python3
"""
Convert "View My Courses" .xlsx exports from Workday-UBC into
RFC 5545 .ics files for Google Calendar.
"""

from __future__ import annotations
import sys, uuid, re, json
from datetime import datetime, timedelta

import pandas as pd
from dateutil import parser as date_parser

from http.server import BaseHTTPRequestHandler
import json
from urllib.parse import parse_qs
import cgi
import tempfile

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
        with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp_file:
            tmp_file.write(file_content)
            tmp_file_path = tmp_file.name
        
        raw = pd.read_excel(tmp_file_path, header=None)
        mask = raw.apply(lambda r: r.astype(str).str.contains(r"Course Listing", case=False, na=False).any(), axis=1)
        hdr_idx = int(mask.idxmax()) if mask.any() else 2
        df = pd.read_excel(tmp_file_path, header=hdr_idx)
        
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
        return None, str(e)


# This is the Vercel Serverless Function handler
class handler(BaseHTTPRequestHandler):
    def do_POST(self):
        try:
            content_type, pdict = cgi.parse_header(self.headers.get('content-type'))
            pdict['boundary'] = bytes(pdict['boundary'], "utf-8")
            
            if content_type == 'multipart/form-data':
                fields = cgi.parse_multipart(self.rfile, pdict)
                file_content = fields.get('file')[0]
                
                ics_data, error = convert_excel_to_ics(file_content)

                if error:
                    self.send_response(400)
                    self.send_header('Content-type', 'application/json')
                    self.end_headers()
                    self.wfile.write(json.dumps({'error': error}).encode('utf-8'))
                else:
                    self.send_response(200)
                    self.send_header('Content-type', 'text/calendar')
                    self.send_header('Content-Disposition', 'attachment; filename="courses.ics"')
                    self.end_headers()
                    self.wfile.write(ics_data.encode('utf-8'))
            else:
                self.send_response(400)
                self.send_header('Content-type', 'application/json')
                self.end_headers()
                self.wfile.write(json.dumps({'error': 'Invalid content type'}).encode('utf-8'))
        
        except Exception as e:
            self.send_response(500)
            self.send_header('Content-type', 'application/json')
            self.end_headers()
            self.wfile.write(json.dumps({'error': str(e)}).encode('utf-8'))

        return

#!/usr/bin/env python3
"""
Convert “View My Courses” .xlsx exports from Workday-UBC into
RFC 5545 .ics files for Google Calendar.
"""

from __future__ import annotations
import sys, uuid, re, json
from datetime import datetime, timedelta

import pandas as pd
from dateutil import parser as date_parser

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

# Break dates to skip (if you use --skip-breaks)
SKIP_DATES = {
    *pd.date_range("2025-12-21", "2026-01-01").date,
    *pd.date_range("2026-02-16", "2026-02-20").date,
}

DAY_MAP = {
    "Mon": "MO", "Tue": "TU", "Wed": "WE", "Thu": "TH",
    "Fri": "FR", "Sat": "SA", "Sun": "SU",
}

def make_uid() -> str:
    return f"{uuid.uuid4()}@ubc-xlsx-to-ics"

def make_dtstamp() -> str:
    return datetime.utcnow().strftime("%Y%m%dT%H%M%SZ")

def parse_meeting_pattern(pat: str):
    try:
        parts = [x.strip() for x in pat.split("|")]
        if len(parts) != 4:
            return (None,) * 6
        # split on hyphen-minus or en-dash, with optional whitespace
        start_date, end_date = re.split(r"\s*[-–]\s*", parts[0])
        days, times, location = parts[1], parts[2], parts[3]
        start_time, end_time = [t.strip() for t in times.split("-")]
        return start_date, end_date, days, start_time, end_time, location
    except Exception:
        return (None,) * 6

def convert_excel_to_ics(path: str, skip_breaks: bool = False):
    # 1) read raw to auto-detect header row
    raw = pd.read_excel(path, header=None)
    hdr_idx = raw.apply(lambda row: "Course Listing" in row.values, axis=1).idxmax()
    df = pd.read_excel(path, header=hdr_idx)

    events = []
    count = 0

    for _, row in df.iterrows():
        course = str(row.get("Course Listing", "")).strip()
        section_raw = str(row.get("Section", "")).strip()
        patterns = str(row.get("Meeting Patterns", "")).strip()
        instr = str(row.get("Instructor", "")).strip()
        if not (course and patterns):
            continue

        # pull out e.g. "211-T1B" from "CPEN_V 211-T1B - Computing Systems I"
        sec_code = None
        if section_raw:
            parts = section_raw.split()
            if len(parts) > 1:
                sec_code = parts[1]
        summary = f"{course} ({sec_code})" if sec_code else course

        for line in patterns.split("\n"):
            sd, ed, days, st, et, loc = parse_meeting_pattern(line)
            if not all([sd, ed, days, st, et]):
                continue
            # map days
            byday = [DAY_MAP[d] for d in DAY_MAP if d in days]
            if not byday:
                continue

            dt_start0 = date_parser.parse(f"{sd} {st}")
            dt_end0   = date_parser.parse(f"{sd} {et}")
            until_dt  = date_parser.parse(f"{ed} {et}")

            if skip_breaks:
                cur = dt_start0.date()
                wk_nums = [list(DAY_MAP.values()).index(d) for d in byday]
                while cur <= until_dt.date():
                    if cur.weekday() in wk_nums and cur not in SKIP_DATES:
                        dts = cur.strftime("%Y%m%d") + dt_start0.strftime("T%H%M%S")
                        dte = cur.strftime("%Y%m%d") + dt_end0.strftime("T%H%M%S")
                        events.append(f"""BEGIN:VEVENT
UID:{make_uid()}
DTSTAMP:{make_dtstamp()}
SUMMARY:{summary}
DTSTART;TZID=America/Vancouver:{dts}
DTEND;TZID=America/Vancouver:{dte}
LOCATION:{loc}
DESCRIPTION:Instructor: {instr}\\nTime: {days} {st}-{et}
END:VEVENT""")
                        count += 1
                    cur += timedelta(days=1)
            else:
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
LOCATION:{loc}
DESCRIPTION:Instructor: {instr}\\nTime: {days} {st}-{et}
END:VEVENT""")
                count += 1

    if count == 0:
        return None, 0

    cal = "\n".join([
        "BEGIN:VCALENDAR",
        "VERSION:2.0",
        "CALSCALE:GREGORIAN",
        VTIMEZONE_BLOCK,
        *events,
        "END:VCALENDAR",
    ])
    return cal, count

if __name__ == "__main__":
    import argparse
    p = argparse.ArgumentParser()
    p.add_argument("file", help="View_My_Courses.xlsx")
    p.add_argument(
        "--skip-breaks", action="store_true",
        help="Remove UBC Winter/Reading Break dates"
    )
    args = p.parse_args()

    try:
        ics, n = convert_excel_to_ics(args.file, skip_breaks=args.skip_breaks)
        if not ics:
            print(json.dumps({
                "error": "No valid course events found. "
                         "Be sure you pointed at a Workday “View My Courses” export."
            }))
            sys.exit(1)
        print(ics)
    except Exception as e:
        print(json.dumps({"error": str(e)}))
        sys.exit(1)

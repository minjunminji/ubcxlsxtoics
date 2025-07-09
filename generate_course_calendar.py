
#!/usr/bin/env python3
"""
Generate a Google‑Calendar‑friendly .ics file from UBC Workday’s **View My Courses** export.

Usage
-----

    python generate_course_calendar.py "View My Courses (1).xlsx" -o courses.ics

Dependencies
------------
    pip install pandas python-dateutil ics pytz
"""

import argparse
import re
from datetime import datetime, timedelta
from pathlib import Path

import pandas as pd
from dateutil import parser as dtparse
from ics import Calendar, Event
import pytz

DAY_MAP = {
    'Mon': 0, 'Tue': 1, 'Wed': 2, 'Thu': 3,
    'Fri': 4, 'Sat': 5, 'Sun': 6
}
ICAL_DAY = {v: k[:2].upper() for k, v in DAY_MAP.items()}  # 0 -> MO, etc.

def clean_header(df: pd.DataFrame) -> pd.DataFrame:
    """Slice off the two pre‑header rows and rename columns."""
    core = df.iloc[2:].copy()
    core.columns = df.iloc[1]
    # Drop entirely blank columns that sometimes appear
    core = core.loc[:, ~core.columns.isna()]
    return core

def parse_meeting_pattern(meeting: str):
    """Return (date_start, date_end, weekdays, time_start, time_end, location, every_two_weeks)."""
    parts = [p.strip() for p in meeting.split('|')]
    if len(parts) < 3:
        return None  # cannot parse

    # ---------- Date range ----------
    date_range = parts[0]
    m = re.match(r'(\d{4}-\d{2}-\d{2}) - (\d{4}-\d{2}-\d{2})', date_range)
    if not m:
        return None
    date_start, date_end = map(lambda s: datetime.strptime(s, '%Y-%m-%d').date(), m.groups())

    # ---------- Days ----------
    days_part = parts[1]
    every_two_weeks = 'Alternate weeks' in days_part
    days_tokens = re.sub(r'\(Alternate weeks\)', '', days_part).strip().split()
    weekdays = []
    for t in days_tokens:
        abbr = t[:3].title()
        if abbr in DAY_MAP:
            weekdays.append(DAY_MAP[abbr])

    # ---------- Times ----------
    time_part = parts[2]
    time_m = re.match(r'([\d:.apm ]+)-([\d:.apm ]+)', time_part, flags=re.I)
    if not time_m:
        return None
    def parse_time(tstr):
        # Workday uses '3:30 p.m.'  -> remove periods, ensure am/pm
        cleaned = tstr.lower().replace('.', '').strip()
        return dtparse.parse(cleaned).time()
    time_start, time_end = map(parse_time, time_m.groups())

    # ---------- Location (optional) ----------
    location = parts[-1] if len(parts) >= 4 else ''

    return date_start, date_end, weekdays, time_start, time_end, location, every_two_weeks

def build_calendar(df: pd.DataFrame, tz_name='America/Vancouver') -> Calendar:
    cal = Calendar()
    tz = pytz.timezone(tz_name)

    for _, row in df.iterrows():
        meeting_info = parse_meeting_pattern(row['Meeting Patterns'])
        if not meeting_info:
            continue

        date_start, date_end, weekdays, time_start, time_end, location, alt_weeks = meeting_info
        course_name = str(row['Course Listing']).strip()
        section = str(row['Section']).strip()
        summary = f"{course_name} ({section})"

        # Create weekly events over the span
        current_date = date_start
        week_counter = 0
        while current_date <= date_end:
            if current_date.weekday() in weekdays:
                if alt_weeks and week_counter % 2 == 1:
                    pass  # skip this week for alternate weeks
                else:
                    ev = Event()
                    ev.name = summary
                    ev.begin = tz.localize(datetime.combine(current_date, time_start))
                    ev.end = tz.localize(datetime.combine(current_date, time_end))
                    ev.location = location
                    cal.events.add(ev)
            # move to next day
            current_date += timedelta(days=1)
            if current_date.weekday() == 0:  # Monday starts a new academic week
                week_counter += 1
    return cal

def main():
    parser = argparse.ArgumentParser(description="Convert UBC View My Courses spreadsheet to .ics")
    parser.add_argument('xlsx', type=Path, help='Input Excel file (View My Courses export)')
    parser.add_argument('-o', '--output', type=Path, default='ubc_courses.ics', help='Output .ics path')
    args = parser.parse_args()

    if not args.xlsx.exists():
        parser.error(f"Input file {args.xlsx} does not exist")

    df_raw = pd.read_excel(args.xlsx, sheet_name='View My Courses', engine='openpyxl')
    df = clean_header(df_raw)

    needed_cols = {'Course Listing', 'Section', 'Meeting Patterns'}
    if not needed_cols.issubset(df.columns):
        missing = ', '.join(sorted(needed_cols - set(df.columns)))
        raise RuntimeError(f"Missing expected column(s): {missing}")

    cal = build_calendar(df)

    with args.output.open('w', encoding='utf-8') as f:
        f.writelines(cal.serialize_iter())

    print(f"✅ Calendar saved to {args.output}")

if __name__ == '__main__':
    main()

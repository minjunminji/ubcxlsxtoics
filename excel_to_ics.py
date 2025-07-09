import pandas as pd
from ics import Calendar, Event
import re
from datetime import datetime, timedelta
from dateutil import parser as date_parser

# Read the Excel file, skipping the first two rows
file_path = 'cpen2courses.xlsx'
df = pd.read_excel(file_path, header=2)

cal = Calendar()

def parse_meeting_pattern(pattern):
    # Example: '2025-09-05 - 2025-11-28 | Fri (Alternate weeks) | 4:00 p.m. - 6:00 p.m. | ESB-Floor 1-Room 1013'
    try:
        parts = [p.strip() for p in pattern.split('|')]
        date_range = parts[0]
        days = parts[1] if len(parts) > 1 else ''
        times = parts[2] if len(parts) > 2 else ''
        location = parts[3] if len(parts) > 3 else ''
        
        # Split date range by ' - ' (space-dash-space) instead of just '-'
        date_parts = date_range.split(' - ')
        if len(date_parts) != 2:
            return None, None, '', '', '', ''
        start_date, end_date = date_parts[0].strip(), date_parts[1].strip()
        
        # Split time range by ' - ' (space-dash-space)
        time_parts = times.split(' - ')
        if len(time_parts) != 2:
            return None, None, '', '', '', ''
        start_time, end_time = time_parts[0].strip(), time_parts[1].strip()
        
        # Remove "(Alternate weeks)" or similar text from days
        days = re.sub(r'\s*\([^)]*\)', '', days).strip()
        
        return start_date, end_date, days, start_time, end_time, location
    except Exception as e:
        print(f'Failed to parse meeting pattern: {pattern} | Error: {e}')
        return None, None, '', '', '', ''

for idx, row in df.iterrows():
    course = str(row.get('Course Listing', '')).strip()
    meeting_patterns = str(row.get('Meeting Patterns', '')).strip()
    instructor = str(row.get('Instructor', '')).strip()
    print(f'Row {idx}:')
    print(f'  Course: {course}')
    print(f'  Meeting Patterns: {meeting_patterns}')
    print(f'  Instructor: {instructor}')
    if not course or not meeting_patterns:
        print('  Skipped: Missing course or meeting pattern')
        continue
    for entry in meeting_patterns.split('\n'):
        entry = entry.strip()
        if not entry:
            continue
        start_date, end_date, days, start_time, end_time, location = parse_meeting_pattern(entry)
        print(f'    Entry: {entry}')
        print(f'      Parsed: start_date={start_date}, end_date={end_date}, days={days}, start_time={start_time}, end_time={end_time}, location={location}')
        if not start_date or not end_date or not days or not start_time or not end_time:
            print('      Skipped: Missing parsed fields')
            continue
        # Parse days (e.g., 'Mon Wed Fri')
        day_map = {'Mon': 'MO', 'Tue': 'TU', 'Wed': 'WE', 'Thu': 'TH', 'Fri': 'FR', 'Sat': 'SA', 'Sun': 'SU'}
        byday = []
        for d in day_map:
            if d in days:
                byday.append(day_map[d])
        print(f'      BYDAY: {byday}')
        if not byday:
            print('      Skipped: No valid days')
            continue
        # Parse start/end datetime using dateutil.parser
        try:
            start_dt = date_parser.parse(start_date + ' ' + start_time)
            end_dt = date_parser.parse(start_date + ' ' + end_time)
            until_dt = date_parser.parse(end_date + ' ' + end_time)
        except Exception as e:
            print(f'      Skipped: Datetime parse error: {e}')
            print(f'      Debug - start_time: "{start_time}", end_time: "{end_time}"')
            continue
        event = Event()
        event.name = course
        event.begin = start_dt
        event.end = end_dt
        event.location = location
        event.description = f'Instructor: {instructor}\nTime: {days} {start_time}-{end_time}'
        # Add recurrence rule (weekly for all courses, including biweekly ones)
        from ics.grammar.parse import Container
        from ics.grammar.parse import ContentLine
        rrule = ContentLine(name='RRULE', value=f'FREQ=WEEKLY;BYDAY={" ,".join(byday)};UNTIL={until_dt.strftime("%Y%m%dT%H%M%SZ")}')
        event.extra.append(rrule)
        cal.events.add(event)
        print(f'      Event added: {course} at {start_dt}')

# Write to ICS file
with open('courses.ics', 'w') as f:
    f.writelines(cal)

# Log the script creation
with open('conversion_log.txt', 'a') as log:
    log.write('Created excel_to_ics.py with fixed parser and weekly recurrence for all courses.\n')

print(f'\nGenerated {len(cal.events)} events in courses.ics')
print(f'Calendar events: {list(cal.events)}') 
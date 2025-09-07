#!/usr/bin/python
# -*- coding: utf-8 -*-

import importlib.util
if importlib.util.find_spec('icalendar') is None:
    raise ImportError('Install the missing icalendar module using "pip install icalendar".')
if importlib.util.find_spec('pymupdf') is None:
    raise ImportError('Install the missing pymupdf module using "pip install pymupdf".')

import argparse
from datetime import datetime, timedelta
import glob
import icalendar # pip install icalendar
import io
from operator import itemgetter
import os.path
import pathlib
import pymupdf # pip install pymupdf
import re
import subprocess
import sys
import tempfile
import time
from urllib import request as urlrequest
from zoneinfo import ZoneInfo

DEFAULT_INPUT = "https://sp3eder.github.io/autosesemenyek/"
TIMEZONE = ZoneInfo("Europe/Budapest")
IMAGE_DPI = 250
HTML_TEMPLATE = """<!DOCTYPE html>
<html lang="hu">
<head>
  <title>Autós Események</title>
  <meta name="color-scheme" content="only light">
  <style>
    @page {
      margin: 0mm;
      size: 210mm 373.3mm;
    }
    html {
      font-family: Bahnschrift, Verdana, Arial, sans-serif;
      font-size: 4mm;
      background-image: url('@SRC_URL@/pattern.jpg');
      background-repeat: repeat;
      background-size: 210mm;
      text-shadow: 0 0 10px white, 0 0 1px #000000b3;
    }
    body {
      margin: 0;
    }
    a {
      color: inherit;
      text-decoration: none;
    }
    table {
      border: none;
      border-collapse: collapse;
      border-spacing: 0;
      table-layout: fixed;
      width: 100%;
    }
    thead {
      background-color: #a7b6d662;
      background: linear-gradient(0deg,#0725384d 0%, #001a4d36 24%, #091b4b1a 79%, #092f4f29 91%, #0a325a17 100%);
      box-shadow: 0px 0px 5px #000000;
      font-family: Verdana, Arial, sans-serif;
      font-weight: bold;
    }
    thead h1 {
      font-size: 1.75em;
      margin: 0;
      margin-top: 6px;
      margin-bottom: 6px;
    }
    thead p {
      font-size: 0.8em;
      font-weight: normal;
      margin: 0;
    }
    thead p:last-child {
      margin-bottom: 4px;
    }
    thead p a {
      font-weight: bold;
    }
    th {
      border: none;
    }
    td {
      border: 1px solid gray;
      border-left: none;
      border-right: none;
      padding: 2px 6px 2px 6px;
    }
    tr td:nth-child(1), tr td:nth-child(2) {
      page-break-inside: avoid !important;
      white-space: nowrap;
    }
  </style>
</head>
<body>
  <table cellspacing="0" cellpadding="4">
    <colgroup>
      <col span="2" style="width:32mm;" /> <!-- Change if font is changed! -->
      <col span="1" style="width:60%;" />
      <col span="1" style="width:40%;" />
    </colgroup>
    <thead>
      <tr><th colspan="4">
        <h1>AUTÓS ESEMÉNYEK NAPTÁRA</h1>
        <p><a href="https://sp3eder.github.io/autosesemenyek/" target="_blank">sp3eder.github.io/autosesemenyek</a> &#8212; eseményleírások, élő követés</p>
        <p class="small"><a href="https://sp3eder.github.io/" target="_blank">sp3eder.github.io</a> &#8212; Autós Appok: alkalmazások az autós közösségnek</p>
      </th></tr>
      <tr><th>KEZDÉS</th><th>VÉGE</th><th>ESEMÉNY</th><th>HELYSZÍN</th></tr>
    </thead>
    <tbody>
      @TABLE_ROWS@
    </tbody>
  </table>
</body>
</html>
"""

def load_html(file_path):
    """Load Autos Esemenyek index.html file from a URL or local file."""
    if re.match(r"^(https?|file)://", file_path):
        with urlrequest.urlopen(file_path) as response:
            return response.read().decode("utf-8")

    with io.open(file_path, mode="r", encoding="utf-8") as file:
        return file.read()

def parse_calids_from_html(html):
    """Extract calendar IDs and colors from the autos esemenyek index.html."""
    pattern = re.compile(r'''{\s*['\"]id['\"]\s*:\s*['\"](?P<calid>[^'\"]+)['\"]\s*,\s*['\"]clr['\"]\s*:\s*['\"](?P<clr>#[0-9A-Fa-f]+)['\"]\s*}''')
    return (m.groupdict() for m in pattern.finditer(html))

def download_calendars(caldata):
    """Download iCal data from Google Calendar public URLs."""
    for [calid, clr] in (itemgetter('calid', 'clr')(cal) for cal in caldata):
        with urlrequest.urlopen(f'https://calendar.google.com/calendar/ical/{calid}/public/basic.ics') as response:
            assert response.status == 200, f"Failed to download calendar {calid}: HTTP {response.status}"
            ics = response.read().decode('utf-8')
            yield { 'ics': ics, 'clr': clr }

def get_calendar_events(caldata):
    """Parse iCal data and extract needed event information."""
    for [ics, clr] in (itemgetter('ics', 'clr')(cal) for cal in caldata):
        cal = icalendar.Calendar.from_ical(ics)
        for component in cal.walk():
            if component.name == "VEVENT":
                yield {
                    'summary': component.get('summary'),
                    'location': component.get('location'),
                    'start': component.start,
                    'end': component.end,
                    'clr': clr
                }

def get_time(dt):
    """Convert date or datetime to datetime, for comparisons."""
    if type(dt) == datetime:
        return dt
    else:
        return datetime(dt.year, dt.month, dt.day, tzinfo=TIMEZONE)

def get_future_events(evtdata, start_of_day):
    """Filter events to only those that are in the future."""
    now = datetime.now(TIMEZONE)
    if start_of_day:
        now = datetime(now.year, now.month, now.day, tzinfo=TIMEZONE)
    return (evt for evt in evtdata if now < get_time(evt['end']))

def events_to_html_table(events):
    """Convert event data to an HTML table."""
    DATE_FMT = '%Y.%m.%d.'
    DATETIME_FMT = DATE_FMT + ' %H:%M'
    def format_dt(dt, end=False):
        # Dates are an open ended interval, so show the day before as the end date
        if end and not isinstance(dt, datetime):
            dt = dt - timedelta(days=1)
        return (dt.astimezone(TIMEZONE).strftime(DATETIME_FMT) if isinstance(dt, datetime)
                else dt.strftime(DATE_FMT))

    html = []
    for evt in events:
        summary = evt.get('summary', '')
        location = evt.get('location', '') or ''
        html.append(
            f'<tr style="color: {evt['clr']};">'
            f'<td>{format_dt(evt['start'])}</td><td>{format_dt(evt['end'], end=True)}</td>'
            f'<td>{summary}</td><td>{location}</td></tr>'
        )

    script_url = pathlib.Path(os.path.dirname(os.path.abspath(__file__))).as_uri()

    return HTML_TEMPLATE.replace('@TABLE_ROWS@', '\n'.join(html), 1).replace('@SRC_URL@', script_url)

def write_pdf_from_html(html):
    """Write HTML to a PDF file using headless Edge browser (Windows only)."""
    output_path = tempfile.mktemp(prefix="event_", suffix=".pdf")

    with tempfile.NamedTemporaryFile(
        mode='w+', encoding='utf-8', prefix='event_', suffix='.html',
        delete=True, delete_on_close=False
    ) as html_file:
        html_path = html_file.name
        html_file.write(html)
        html_file.close()

        subprocess.run([
            'C:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedge.exe',
            '--headless',
            '--disable-gpu',
            '--run-all-compositor-stages-before-draw',
            '--no-pdf-header-footer',
            '--print-to-pdf-no-header',
            f'--print-to-pdf={output_path}',
            html_path
        ], check=True, stdout=sys.stdout, stderr=sys.stderr, text=True)

        timeout = 10
        start_time = time.time()
        while not os.path.exists(output_path):
            if time.time() - start_time > timeout:
                raise TimeoutError(f'PDF file was not created within {timeout} seconds: {output_path}')
            time.sleep(0.1)

    return output_path
    
def export_to_png(output_path, pdf_path):
    """Convert PDF pages to PNG images using PyMuPDF (if available)."""
    if not os.path.isabs(output_path):
        output_path = os.path.join(os.getcwd(), output_path)

    for file in glob.glob(f"{output_path}_*.png"):
        os.remove(file)

    with pymupdf.open(pdf_path) as doc:
        for i, page in enumerate(doc):
            pix = page.get_pixmap(dpi=IMAGE_DPI)
            pix.save(f"{output_path}_{i+1}.png")

    try:
        os.remove(pdf_path)
    except OSError:
        pass

    return os.path.dirname(output_path)

def main():
    """Main function to convert calendars to PDF."""
    if os.name != "nt":
        raise OSError("This script can only be run on Windows.")

    parser = argparse.ArgumentParser(description="Export calendar events as a PDF.")
    parser.add_argument("html_file", nargs="?", default=DEFAULT_INPUT, help="URL or path of HTML file to parse. Default: website.")
    parser.add_argument("--output", "-o", default="events", help="Output without extension. Default: events")
    parser.add_argument("--start-of-day", type=bool, default=False, help="Show events from midnight today. Default: False")
    args = parser.parse_args()

    print("Loading HTML...")
    caldata = parse_calids_from_html(load_html(args.html_file))
    print(f"Loaded some calendars. Processing...")
    caldata = download_calendars(caldata)
    evtdata = get_calendar_events(caldata)
    evtdata = get_future_events(evtdata, args.start_of_day)
    events = sorted(evtdata, key=lambda evt: (get_time(evt["start"]), evt["summary"]))
    print(f"Found {len(events)} future events. Generating PDF...")
    html = events_to_html_table(events)
    pdf_file = write_pdf_from_html(html)
    print(f"PDF file is created. Converting to PNG images...")
    out_dir = export_to_png(args.output, pdf_file)
    os.startfile(out_dir)

if __name__=="__main__":
    main()

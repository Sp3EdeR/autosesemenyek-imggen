import importlib.util
if importlib.util.find_spec('icalendar') is None:
    raise ImportError('Install the missing icalendar module using "pip install icalendar".')

import argparse
from datetime import datetime
import errno
import icalendar # pip install icalendar
import io
from operator import itemgetter
import os.path
import pathlib
import re
import subprocess
import sys
import tempfile
import time
from urllib import request as urlrequest
from zoneinfo import ZoneInfo

DEFAULT_FILE = "https://sp3eder.github.io/autosesemenyek/"
TIMEZONE = ZoneInfo("Europe/Budapest")
AD_FREQUENCY = 15 # Insert ad at every Nth event
HTML_TEMPLATE = '''<!DOCTYPE html>
<html lang="hu">
<head>
  <meta name="color-scheme" content="only light">
  <script src="https://cdn.jsdelivr.net/npm/jquery@3.7.1/dist/jquery.min.js" integrity="sha256-/JqT3SQfawRcv/BIHPThkBvs0OEvtFFmqPF/lYI/Cxo=" crossorigin="anonymous" referrerpolicy="no-referrer"></script>
  <script>
    $('document').ready(function() {
      $('tr.ad i').each(function() {
        $(this).replaceWith(
          '<h3>AUTÓS ESEMÉNYEK NAPTÁRA</h3>' +
          'további infók &#8212; <a href="https://sp3eder.github.io/autosesemenyek/" target="_blank">https://sp3eder.github.io/autosesemenyek/</a> &#8212; élő követés'
        );
      });
    });
  </script>
  <style>
    @page {
      margin: 0mm;
      size: 210mm 373.3mm;
    }
    body {
      font-family: Arial, sans-serif;
      font-size: 16px;
      background-image: url('@SRC_URL@/pattern.jpg');
      background-repeat: repeat;
      background-size: 80%;
      margin: 0;
      text-shadow: 0 0 10px white, 0 0 1px #000000b3;
    }
    a {
      color: inherit;
      text-decoration: none;
    }
    table {
      border: 1px solid gray;
      border-collapse: collapse;
      border-spacing: 0;
      width: 100%;
    }
    thead {
      background-color: #afbdcec5;
      font-weight: bold;
    }
    tr td:nth-child(1), tr td:nth-child(2) {
      page-break-inside: avoid !important;
      white-space: nowrap;
    }
    td, th {
      border: 1px solid gray;
      padding: 2px 6px 2px 6px;
    }
    tr.ad {
      background-color: #cdd0d631;
      text-align: center;
    }
    tr.ad h3 {
      margin: 0;
      font-size: 1.2em;
    }
  </style>
</head>
<body>
  <table cellspacing="0" cellpadding="4">
    <thead><tr><th>Kezd\u00e9s</th><th>V\u00e9ge</th><th>Esem\u00e9ny</th><th>Helysz\u00edn</th></tr></thead>
    <tbody>
      @TABLE_ROWS@
    </tbody>
  </table>
</body>
</html>
'''

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
    AD_ROW = f'<tr class="ad"><td colspan="4"><i></i></td></tr>'
    DATE_FMT = '%Y.%m.%d.'
    DATETIME_FMT = DATE_FMT + ' %H:%M'
    def format_dt(dt):
        return (dt.astimezone(TIMEZONE).strftime(DATETIME_FMT) if isinstance(dt, datetime)
                else dt.strftime(DATE_FMT))

    html = []
    for idx, evt in enumerate(events):
        if 0 < idx and idx % AD_FREQUENCY == 0:
            html.append(AD_ROW)

        summary = evt.get('summary', '')
        location = evt.get('location', '') or ''
        html.append(
            f'<tr style="color: {evt['clr']};">'
            f'<td>{format_dt(evt['start'])}</td><td>{format_dt(evt['end'])}</td>'
            f'<td>{summary}</td><td>{location}</td></tr>'
        )
    html.append(AD_ROW)

    script_url = pathlib.Path(os.path.dirname(os.path.abspath(__file__))).as_uri()

    return HTML_TEMPLATE.replace('@TABLE_ROWS@', '\n'.join(html), 1).replace('@SRC_URL@', script_url)

def write_pdf_from_html(output_path, html):
    """Write HTML to a PDF file using headless Edge browser (Windows only)."""
    if not os.path.isabs(output_path):
        output_path = os.path.join(os.getcwd(), output_path)
    try:
        os.remove(output_path)
    except OSError as e:
        if e.errno != errno.ENOENT:
            raise

    with tempfile.NamedTemporaryFile(
        mode='w+', encoding='utf-8', suffix='.html', delete=True, delete_on_close=False
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
                raise TimeoutError(f"PDF file was not created within {timeout} seconds: {output_path}")
            time.sleep(0.1)

        return output_path
    
def open_file(filepath):
    """Open a file with the default associated application (Windows only)."""
    subprocess.Popen(
        [filepath], shell=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
        creationflags=subprocess.DETACHED_PROCESS | subprocess.CREATE_NEW_PROCESS_GROUP
    )

def main():
    """Main function to convert calendars to PDF."""
    if os.name != "nt":
        raise OSError("This script can only be run on Windows.")
    
    parser = argparse.ArgumentParser(description="Export calendar events as a PDF.")
    parser.add_argument("html_file", nargs="?", default=DEFAULT_FILE, help="URL or path of HTML file to parse. Default: website.")
    parser.add_argument("--output", "-o", default="events.pdf", help="Output PDF file. Default: events.pdf")
    parser.add_argument("--start-of-day", type=bool, default=False, help="Show events from midnight today. Default: False")
    args = parser.parse_args()

    print("Loading HTML...")
    caldata = parse_calids_from_html(load_html(args.html_file))
    print(f"Loaded some calendars. Processing...")
    caldata = download_calendars(caldata)
    evtdata = get_calendar_events(caldata)
    evtdata = get_future_events(evtdata, args.start_of_day)
    events = sorted(evtdata, key=lambda evt: (get_time(evt['start']), evt['summary']))
    print(f"Found {len(events)} future events. Generating PDF...")
    html = events_to_html_table(events)
    output = write_pdf_from_html(args.output, html)
    print(f"Done. Output written to {output}")
    open_file(output)

if __name__=="__main__":
    main()

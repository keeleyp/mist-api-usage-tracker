#!/usr/bin/env python3
"""Track how many Mist API calls are left on each cloud over the day.

On start-up, then at hh:09, hh:19, ... hh:59 (configurable), the script calls
GET /api/v1/self/usage on each cloud listed in mist_api_usage_tracker.ini. It
adds a row to the spreadsheet (the time, then the calls still available,
request_limit - requests, one column per cloud) and prints the results as a
table in the terminal. The check itself costs one call per cloud per sample.

Runs until stopped with Ctrl+C. Each row is saved as soon as it is written, so
the spreadsheet is always up to date. If the file already exists, new rows are
added to the end of it.
"""
import configparser
import os
import sys
import time
from datetime import datetime
from urllib.parse import urlparse

import requests
from openpyxl import Workbook, load_workbook
from openpyxl.chart import LineChart, Reference
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
INI_NAME = "mist_api_usage_tracker.ini"
config = configparser.ConfigParser()
config.optionxform = str  # keep cloud names exactly as written
if not config.read(os.path.join(SCRIPT_DIR, INI_NAME)):
    sys.exit(f"{INI_NAME} not found next to the script - copy {INI_NAME}.example and fill it in.")

INTERVAL_MINUTES = config.getint("settings", "interval_minutes", fallback=10)
# Minutes past each interval boundary to sample at: 9 with a 10-minute
# interval gives hh:09, hh:19, hh:29, hh:39, hh:49, hh:59.
OFFSET_MINUTES = config.getint("settings", "offset_minutes", fallback=9)
OUTPUT_FILE = os.path.join(SCRIPT_DIR, os.path.expanduser(
    config.get("settings", "output_file", fallback="mist_api_usage.xlsx")))

REQUEST_TIMEOUT = 30
DATA_SHEET = "Usage"
CHART_SHEET = "Chart"
TIME_HEADER = "Time"
NOTES_HEADER = "Notes"
HEADER_FONT = Font(bold=True, color="FFFFFF")
HEADER_FILL = PatternFill("solid", fgColor="1F4E78")


def build_session():
    """Shared HTTP session, wired up for a TLS-inspecting proxy (e.g. Zscaler)
    when the optional [network] section is present in the .ini. With no
    [network] section, behaviour is unchanged (certifi CA store, and the
    HTTP_PROXY / HTTPS_PROXY / NO_PROXY environment variables if set)."""
    session = requests.Session()
    if "network" not in config:
        return session

    section = config["network"]
    ca_bundle = section.get("ca_bundle", "").strip()
    if ca_bundle:
        ca_path = os.path.expanduser(ca_bundle)
        if not os.path.isfile(ca_path):
            sys.exit(f"[network] ca_bundle does not exist: {ca_path}\n"
                     f"This should be a PEM file containing your proxy's root CA "
                     f"certificate (e.g. exported from Zscaler).")
        session.verify = ca_path
    elif not section.getboolean("verify_ssl", fallback=True):
        session.verify = False
        print("WARNING: TLS certificate verification is DISABLED ([network] verify_ssl = false). "
              "Only use this as a last resort.")
        requests.packages.urllib3.disable_warnings(requests.packages.urllib3.exceptions.InsecureRequestWarning)

    for scheme in ("http", "https"):
        proxy = section.get(f"{scheme}_proxy", "").strip()
        if proxy:
            session.proxies[scheme] = proxy
    return session


SESSION = build_session()


def cloud_to_host(cloud):
    """Normalise an .ini cloud value to an API host name like api.eu.mist.com."""
    value = cloud.lower().strip()
    if "://" in value:
        value = urlparse(value).netloc
    value = value.strip("/")
    if value.endswith("mist.com"):
        if value.startswith("manage."):
            value = "api." + value[len("manage."):]
        return value
    if value in ("us", "global", "global01"):
        return "api.mist.com"
    return f"api.{value}.mist.com"


def load_clouds():
    """Return [(api_host, token)] from the [clouds] section, in file order."""
    if "clouds" not in config or not config["clouds"]:
        sys.exit(f"No clouds listed - add 'cloud = api_token' lines under [clouds] in {INI_NAME}.")
    clouds = []
    for cloud, token in config["clouds"].items():
        token = token.strip()
        if not token or token.startswith("YOUR_"):
            sys.exit(f"No API token set for '{cloud}' in {INI_NAME}.")
        host = cloud_to_host(cloud)
        if any(host == h for h, _ in clouds):
            sys.exit(f"'{cloud}' is listed more than once in {INI_NAME} (as {host}).")
        clouds.append((host, token))
    return clouds


def fetch_available(host, token):
    """Return (calls_available, request_limit, note). calls_available and
    request_limit are None on failure."""
    url = f"https://{host}/api/v1/self/usage"
    try:
        resp = SESSION.get(url, headers={"Authorization": f"Token {token}"}, timeout=REQUEST_TIMEOUT)
    except requests.exceptions.SSLError as e:
        return None, None, (f"TLS error ({e}) - if you're behind Zscaler or a similar proxy, "
                      f"set [network] ca_bundle in {INI_NAME}")
    except requests.exceptions.ProxyError as e:
        return None, None, f"proxy error ({e}) - check [network] http_proxy / https_proxy in {INI_NAME}"
    except requests.exceptions.RequestException as e:
        return None, None, f"could not connect ({e.__class__.__name__})"

    if resp.status_code == 401:
        return None, None, "HTTP 401 - token not valid on this cloud"
    if resp.status_code == 429:
        return 0, None, "HTTP 429 - rate limited (no calls left)"
    if not resp.ok:
        return None, None, f"HTTP {resp.status_code}"
    data = resp.json()
    limit = data.get("request_limit", 5000)
    return limit - data.get("requests", 0), limit, None


def style_header(cell):
    cell.font = HEADER_FONT
    cell.fill = HEADER_FILL
    cell.alignment = Alignment(horizontal="center")


def open_workbook(hosts):
    """Open (or create) the output workbook and make sure every cloud has a
    column. Returns (workbook, sheet, {header: column_number})."""
    if os.path.exists(OUTPUT_FILE):
        wb = load_workbook(OUTPUT_FILE)
        if DATA_SHEET not in wb.sheetnames:
            sys.exit(f"{OUTPUT_FILE} exists but has no '{DATA_SHEET}' sheet - "
                     f"move it aside or change output_file in {INI_NAME}.")
        ws = wb[DATA_SHEET]
        columns = {ws.cell(1, c).value: c for c in range(1, ws.max_column + 1) if ws.cell(1, c).value}
        print(f"Appending to existing {OUTPUT_FILE}")
    else:
        wb = Workbook()
        ws = wb.active
        ws.title = DATA_SHEET
        columns = {TIME_HEADER: 1}
        style_header(ws.cell(1, 1, TIME_HEADER))
        ws.column_dimensions["A"].width = 20
        ws.freeze_panes = "B2"
        print(f"Creating {OUTPUT_FILE}")

    # Clouds added to the .ini since the file was created get new columns,
    # placed before Notes so the cloud columns stay together for the chart.
    notes_col = columns.pop(NOTES_HEADER, None)
    for host in hosts:
        if host not in columns:
            col = max(columns.values()) + 1
            if notes_col and col >= notes_col:
                ws.insert_cols(notes_col)
                notes_col += 1
            columns[host] = col
            style_header(ws.cell(1, col, host))
            ws.column_dimensions[get_column_letter(col)].width = 20
    if not notes_col:
        notes_col = max(columns.values()) + 1
        style_header(ws.cell(1, notes_col, NOTES_HEADER))
        ws.column_dimensions[get_column_letter(notes_col)].width = 60
    columns[NOTES_HEADER] = notes_col
    return wb, ws, columns


def rebuild_chart(wb, ws, columns):
    """Line chart of calls available over time for every cloud column."""
    if CHART_SHEET in wb.sheetnames:
        del wb[CHART_SHEET]
    if ws.max_row < 2:
        return
    cs = wb.create_sheet(CHART_SHEET)
    chart = LineChart()
    chart.title = "Mist API calls available"
    chart.y_axis.title = "Calls available"
    chart.x_axis.title = TIME_HEADER
    chart.x_axis.number_format = "hh:mm"
    chart.height, chart.width = 12, 28
    cloud_cols = sorted(c for h, c in columns.items() if h not in (TIME_HEADER, NOTES_HEADER))
    for col in cloud_cols:
        chart.add_data(Reference(ws, min_col=col, min_row=1, max_row=ws.max_row), titles_from_data=True)
    chart.set_categories(Reference(ws, min_col=1, min_row=2, max_row=ws.max_row))
    cs.add_chart(chart, "A1")


def save(wb):
    """Save the workbook; on failure keep going so the row is saved next time."""
    try:
        wb.save(OUTPUT_FILE)
        return True
    except PermissionError:
        print(f"  WARNING: could not save {OUTPUT_FILE} (is it open in Excel?) - "
              f"will try again at the next sample.")
        return False


def print_results(now, results, previous, saved):
    """Show one sample as a table: calls left, used this hour, limit, and the
    change in calls left since the previous sample this run."""
    print(f"\n=== {now:%Y-%m-%d %H:%M:%S} " + ("" if saved else "(NOT SAVED) ") + "=" * 40)
    print(f"{'Cloud':<20}{'Calls left':>12}{'Used':>8}{'Limit':>8}{'Change':>9}")
    print("-" * 57)
    for host, available, limit, note in results:
        if available is None:
            print(f"{host:<20}{'ERROR':>12}   {note}")
            continue
        used = f"{limit - available}" if limit is not None else "-"
        limit_s = f"{limit}" if limit is not None else "-"
        change = f"{available - previous[host]:+d}" if previous.get(host) is not None else ""
        line = f"{host:<20}{available:>12,}{used:>8}{limit_s:>8}{change:>9}"
        print(line + (f"   {note}" if note else ""))
    lows = [(a, h) for h, a, _, _ in results if a is not None]
    if lows:
        a, h = min(lows)
        print("-" * 57)
        print(f"Lowest: {h} with {a:,} calls left")


def take_sample(wb, ws, columns, clouds, previous):
    """Query every cloud, write one row, save and print the results. Updates
    `previous` with this sample's values for the next change column."""
    now = datetime.now().replace(microsecond=0)
    row = ws.max_row + 1
    ws.cell(row, columns[TIME_HEADER], now).number_format = "yyyy-mm-dd hh:mm"
    results, notes = [], []
    for host, token in clouds:
        available, limit, note = fetch_available(host, token)
        if available is not None:
            ws.cell(row, columns[host], available)
        if note:
            notes.append(f"{host}: {note}")
        results.append((host, available, limit, note))
    if notes:
        ws.cell(row, columns[NOTES_HEADER], "; ".join(notes))
    rebuild_chart(wb, ws, columns)
    saved = save(wb)
    print_results(now, results, previous, saved)
    previous.update({host: available for host, available, _, _ in results})


def seconds_to_next_sample():
    """Seconds until the next scheduled sample, e.g. 09:09:00, 09:19:00."""
    interval = INTERVAL_MINUTES * 60
    # Local-time alignment so samples land on the right minutes in the user's timezone.
    local_now = time.time() + datetime.now().astimezone().utcoffset().total_seconds()
    return interval - ((local_now - OFFSET_MINUTES * 60) % interval)


def main():
    if INTERVAL_MINUTES < 1:
        sys.exit("interval_minutes must be at least 1.")
    if not 0 <= OFFSET_MINUTES < INTERVAL_MINUTES:
        sys.exit("offset_minutes must be at least 0 and less than interval_minutes.")
    clouds = load_clouds()
    wb, ws, columns = open_workbook([h for h, _ in clouds])
    print(f"Tracking {len(clouds)} cloud(s) every {INTERVAL_MINUTES} min - press Ctrl+C to stop.")
    previous = {}
    try:
        # One entry straight away on start-up, then on the schedule.
        take_sample(wb, ws, columns, clouds, previous)
        while True:
            wait = seconds_to_next_sample()
            print(f"\nNext sample at {datetime.fromtimestamp(time.time() + wait):%H:%M}")
            # +1s so a slightly early wake-up can't trigger a double sample.
            time.sleep(wait + 1)
            take_sample(wb, ws, columns, clouds, previous)
    except KeyboardInterrupt:
        print("\nStopping.")
        save(wb)
        print(f"Saved {OUTPUT_FILE}")


if __name__ == "__main__":
    main()

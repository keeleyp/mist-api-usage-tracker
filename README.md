# Mist API Usage Tracker

Records how many Mist API calls are left on each cloud through the day, so you can see the trend.

When it starts, and then at hh:09, hh:19, hh:29, hh:39, hh:49 and hh:59, the script calls `GET /api/v1/self/usage` on every cloud in the .ini. It then adds a row to `mist_api_usage.xlsx` with the time and the calls left (`request_limit - requests`) for each cloud. Each cloud has its own column.

It also prints each sample as a table in the terminal. The table shows calls left, calls used this hour, the limit, the change since the last sample, and the cloud with the fewest calls left.

Each check uses **1 call per cloud**, so 6 calls per cloud per hour on the default schedule.

## Setup

```bash
pip install -r requirements.txt
cp mist_api_usage_tracker.ini.example mist_api_usage_tracker.ini
```

Edit `mist_api_usage_tracker.ini`. Under `[clouds]`, add one line per cloud in the form `cloud = api_token`.

## Run

```bash
python3 mist_api_usage_tracker.py
```

The script runs until you press Ctrl+C. It saves the file after every sample. If the file already exists, new rows are added to the end of it.

## Output

- **Usage** sheet: Time, then one column per cloud with the calls left, then Notes. Notes records any errors, such as a 401 when a token isn't valid on that cloud. A failed cloud leaves its cell blank.
- **Chart** sheet: a line chart of calls left over time for each cloud.

On Windows, close the workbook in Excel before a sample is due. Excel locks the file while it's open. If a save fails, the row is kept and saved at the next sample.

## Settings

- `interval_minutes`: how often to sample, in minutes. Default 10.
- `offset_minutes`: how many minutes past each interval to sample. Default 9, which gives hh:09, hh:19 and so on.
- `output_file`: name of the spreadsheet.
- Optional `[network]` section for Zscaler or other proxies: `ca_bundle`, `verify_ssl`, `http_proxy`, `https_proxy`.

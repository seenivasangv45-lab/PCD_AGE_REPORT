# Monthly Report Dashboard — Web Application

A professional web-based dashboard that converts your PCD_20 + Age_24 data into interactive monthly reports. Multiple users on the same WiFi network can access it simultaneously.

## Quick Start

### Windows
1. Double-click **`run_server.bat`**
2. Browser opens automatically to `http://localhost:5000`
3. Share the **Network URL** (shown in terminal) with other users

### Mac / Linux
1. Double-click **`run_server.sh`** (or run `bash run_server.sh`)
2. Browser opens automatically
3. Share the Network URL with other users

## Requirements

- **Python 3.8+** — [download here](https://python.org)
- Dependencies auto-install on first run: Flask, pandas, openpyxl, numpy

## How It Works

1. **Upload two files separately** — PCD_20 and Age_24, each as **CSV or XLSX**
2. Files are **auto-detected** by filename (e.g. `pcd_20_march.csv`, `age_24.xlsx`) or by column structure
3. If you drop a file in the wrong slot, it auto-corrects to the right one
4. Select your **date range** (start month → end month) — filters by `Svc_Date`
5. Click **Generate Dashboard Report**
6. View the interactive dashboard or download the styled Excel file

## File Format Support

| Format | PCD_20 | Age_24 |
|--------|--------|--------|
| .xlsx  | ✓      | ✓      |
| .csv   | ✓      | ✓      |

**Auto-detection priority:**
1. Filename contains `pcd` or `age` (case-insensitive)
2. Column structure matching (Svc_Date + Crg_Amt → PCD; Financial_Class + textbox18 → Age)

## Network Access

When the server starts, it displays two URLs:
- `http://localhost:5000` — for the machine running the server
- `http://192.168.x.x:5000` — for other PCs on the same WiFi

Share the second URL with your team members.

## File Structure

```
report_app/
├── app.py              # Flask server + processing logic
├── run_server.bat      # Windows double-click launcher
├── run_server.sh       # Mac/Linux double-click launcher
├── static/
│   └── index.html      # Professional web UI
├── uploads/            # (auto-created) Temp uploaded files
├── outputs/            # (auto-created) Generated reports
└── README.md
```

#!/usr/bin/env python3
"""
Automated Data Synchronizer for Godrej Training & Merchandising Dashboards
Downloads latest live datasets directly from SharePoint public export URLs,
embeds them into index.html, training.html, and m_score.html, and automatically
pushes changes to GitHub.
"""

import urllib.request
import http.cookiejar
from pyxlsb import open_workbook
import openpyxl
import io
import json
import datetime
import functools
import re
import os
import subprocess
import sys
import time

print = functools.partial(print, flush=True)

SHAREPOINT_URLS = {
    "index": "https://teamchannelplay-my.sharepoint.com/:x:/g/personal/bikash_roy1_channelplay_in/IQBx5HIst0LPT4_moEtMpsbtAd4w3ClOl0h-mrlnCEmDCno?download=1",
    "training": "https://teamchannelplay-my.sharepoint.com/:x:/g/personal/bikash_roy1_channelplay_in/IQBRmCEH6nI8TLFuK-RVqPu0ATXbidF7rGfITZvpZH7PyAA?download=1",
    "m_score": "https://teamchannelplay-my.sharepoint.com/:x:/g/personal/bikash_roy1_channelplay_in/IQALwrJKNKLQTIXnzWX8IQrXAUJP1FnSVdMZD7jPMrqpxvo?e=yACydV&download=1",
    "program_performance": "https://teamchannelplay-my.sharepoint.com/:x:/g/personal/bikash_roy1_channelplay_in/IQB-EkkYdgTFQphMhXNUEfKSAWIldx1iKI_TWThpQI42w8E?e=2dQVkl&download=1",
    "branch": "https://teamchannelplay-my.sharepoint.com/:x:/g/personal/bikash_roy1_channelplay_in/IQB-EkkYdgTFQphMhXNUEfKSAWIldx1iKI_TWThpQI42w8E?e=2dQVkl&download=1"
}

ONEDRIVE_BASE = "/Users/bikash/Library/CloudStorage/OneDrive-ChannelplayLimited/My Laptop/0 Active Projects/Godrej VM 100838/0 Project 2.0/Reports for dashboard"

LOCAL_FALLBACKS = {
    "index": [
        os.path.join(ONEDRIVE_BASE, "QC Audit Report.xlsx"),
        os.path.join(os.path.dirname(os.path.abspath(__file__)), "QC Audit Report.xlsx"),
        os.path.join(os.path.dirname(os.path.abspath(__file__)), "qc_download.xlsx"),
    ],
    "training": [
        os.path.join(ONEDRIVE_BASE, "Training Tracker.xlsx"),
        os.path.join(os.path.dirname(os.path.abspath(__file__)), "Training Tracker.xlsx"),
    ],
    "m_score": [
        os.path.join(ONEDRIVE_BASE, "Ops Reports", "Godrej MS Value.xlsx"),
        os.path.join(ONEDRIVE_BASE, "Ops Reports", "Godrej MS.xlsb"),
        os.path.join(os.path.dirname(os.path.abspath(__file__)), "Godrej MS Value.xlsx"),
        os.path.join(os.path.dirname(os.path.abspath(__file__)), "Godrej MS.xlsb"),
    ],
    "program_performance": [
        os.path.join(ONEDRIVE_BASE, "Ops Reports", "Godrej VM Productivity Report.xlsb"),
        os.path.join(os.path.dirname(os.path.abspath(__file__)), "Godrej VM Productivity Report.xlsb"),
    ],
    "branch": [
        os.path.join(ONEDRIVE_BASE, "Ops Reports", "Godrej VM Productivity Report.xlsb"),
        os.path.join(os.path.dirname(os.path.abspath(__file__)), "Godrej VM Productivity Report.xlsb"),
    ]
}

def get_opener():
    cookie_jar = http.cookiejar.CookieJar()
    return urllib.request.build_opener(
        urllib.request.HTTPCookieProcessor(cookie_jar),
        urllib.request.HTTPRedirectHandler
    )

def to_direct_sharepoint_url(url):
    m = re.search(r'/personal/[^/]+/([A-Za-z0-9_-]+)', url)
    if not m:
        m = re.search(r'share=([A-Za-z0-9_-]+)', url)
    if m:
        share_id = m.group(1)
        return f"https://teamchannelplay-my.sharepoint.com/personal/bikash_roy1_channelplay_in/_layouts/15/download.aspx?share={share_id}"
    return url

def fetch_url_data(opener, url, fallback_paths=None, timeout=90, retries=3, label="dataset"):
    # When running on local machine where OneDrive files exist and user hasn't requested --remote:
    # Read the active local file directly. This guarantees instant updates with zero SharePoint cache delay!
    prefer_remote = "--remote" in sys.argv or "--prefer-remote" in sys.argv
    if not prefer_remote and fallback_paths:
        for path in fallback_paths:
            if os.path.exists(path) and os.path.getsize(path) > 3000:
                print(f"✓ Using active local OneDrive file for {label}: {path}")
                with open(path, "rb") as f:
                    return f.read()

    print(f"Fetching latest {label} from SharePoint...")
    urls_to_try = [to_direct_sharepoint_url(url), url]
    last_err = None

    for target_url in urls_to_try:
        req = urllib.request.Request(
            target_url,
            headers={"User-Agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36"}
        )
        for attempt in range(1, retries + 1):
            try:
                resp = opener.open(req, timeout=timeout)
                chunks = []
                while True:
                    chunk = resp.read(1024 * 1024)
                    if not chunk:
                        break
                    chunks.append(chunk)
                data = b"".join(chunks)
                if len(data) > 3000 and not (data.startswith(b"<!") or data.startswith(b"<html")):
                    print(f"✓ Downloaded {len(data):,} bytes for {label}")
                    return data
                else:
                    print(f"Warning: Downloaded data for {label} was too small or HTML error page.")
            except Exception as e:
                last_err = e
                if attempt < retries:
                    time.sleep(2)

    if fallback_paths:
        print(f"SharePoint fetch failed for {label}: {last_err}. Checking local fallbacks...")
        for path in fallback_paths:
            if os.path.exists(path):
                print(f"✓ Using local fallback for {label}: {path}")
                with open(path, "rb") as f:
                    return f.read()

    if last_err:
        raise last_err
    raise ValueError(f"Could not download {label} or locate local fallback file.")

def embed_sample_data(html_file, rows):
    if not os.path.exists(html_file):
        print(f"File {html_file} not found, skipping.")
        return
    with open(html_file, "r", encoding="utf-8") as f:
        content = f.read()

    # Minify JSON string safely
    json_str = json.dumps(rows, separators=(",", ":"))
    new_tag = f'<script id="sample-data" type="application/json">{json_str}</script>'
    pattern = r'<script id="sample-data" type="application/json">.*?</script>'

    if re.search(pattern, content, flags=re.DOTALL):
        updated_content = re.sub(pattern, lambda m: new_tag, content, count=1, flags=re.DOTALL)
    else:
        # Fallback if tag doesn't exist yet: insert before </body>
        updated_content = content.replace("</body>", f"{new_tag}\n</body>")

    with open(html_file, "w", encoding="utf-8") as f:
        f.write(updated_content)
    print(f"✓ Embedded {len(rows):,} rows into {html_file} ({len(json_str):,} chars)")

def sync_index(opener):
    print("\n--- Syncing index.html (QC Audit Report) ---")
    data = fetch_url_data(opener, SHAREPOINT_URLS["index"], LOCAL_FALLBACKS.get("index"), label="QC Audit Report (index)")
    wb = openpyxl.load_workbook(io.BytesIO(data), read_only=True, data_only=True)
    ws = wb[wb.sheetnames[0]]
    rows = []
    for row in ws.iter_rows(values_only=True):
        row_clean = []
        for cell in row:
            if isinstance(cell, (datetime.date, datetime.datetime)):
                row_clean.append(cell.strftime("%Y-%m-%d"))
            else:
                row_clean.append(cell)
        rows.append(row_clean)
    embed_sample_data("index.html", rows)

def sync_training(opener):
    print("\n--- Syncing training.html (Training Tracker) ---")
    data = fetch_url_data(opener, SHAREPOINT_URLS["training"], LOCAL_FALLBACKS.get("training"), label="Training Tracker (training)")
    wb = openpyxl.load_workbook(io.BytesIO(data), read_only=True, data_only=True)
    target_sheet = "Training Details" if "Training Details" in wb.sheetnames else ("Training Data" if "Training Data" in wb.sheetnames else wb.sheetnames[0])
    ws = wb[target_sheet]
    rows = []
    for row in ws.iter_rows(values_only=True):
        if not any(c is not None and str(c).strip() != "" for c in row):
            continue
        row_clean = []
        for cell in row:
            if isinstance(cell, (datetime.date, datetime.datetime)):
                row_clean.append(cell.strftime("%Y-%m-%d"))
            else:
                row_clean.append(cell)
        rows.append(row_clean)
    embed_sample_data("training.html", rows)

def sync_m_score(opener):
    print("\n--- Syncing m_score.html (Product-VM-Score) ---")
    data = fetch_url_data(opener, SHAREPOINT_URLS["m_score"], LOCAL_FALLBACKS.get("m_score"), timeout=120, label="Godrej MS (m_score)")
    
    rows = []
    if data[:4] == b"PK\x03\x04":
        wb = openpyxl.load_workbook(io.BytesIO(data), read_only=True, data_only=True)
        sheet_name = "Product-VM-Score" if "Product-VM-Score" in wb.sheetnames else wb.sheetnames[0]
        ws = wb[sheet_name]
        for i, row in enumerate(ws.iter_rows(values_only=True)):
            r_vals = list(row[:30])
            if i > 0 and len(r_vals) > 14:
                d = r_vals[14]
                if isinstance(d, (int, float)):
                    dt = datetime.date(1899, 12, 30) + datetime.timedelta(days=int(d))
                    r_vals[14] = dt.strftime("%Y-%m-%d")
                elif isinstance(d, (datetime.date, datetime.datetime)):
                    r_vals[14] = d.strftime("%Y-%m-%d")
            rows.append(r_vals)
    else:
        with open_workbook(io.BytesIO(data)) as wb:
            sheet_name = "Product-VM-Score" if "Product-VM-Score" in wb.sheets else wb.sheets[0]
            with wb.get_sheet(sheet_name) as s:
                for i, row in enumerate(s.rows()):
                    r_vals = [c.v for c in row[:30]]
                    if i > 0 and len(r_vals) > 14:
                        d = r_vals[14]
                        if isinstance(d, (int, float)):
                            dt = datetime.date(1899, 12, 30) + datetime.timedelta(days=int(d))
                            r_vals[14] = dt.strftime("%Y-%m-%d")
                    rows.append(r_vals)
    embed_sample_data("m_score.html", rows)

def sync_program_performance(opener):
    print("\n--- Syncing program_performance.html (Program Performance) ---")
    data = fetch_url_data(opener, SHAREPOINT_URLS["program_performance"], LOCAL_FALLBACKS.get("program_performance"), timeout=60, label="VM Productivity (program_performance)")
    with open_workbook(io.BytesIO(data)) as wb:
        sheet_name = "Program Performance" if "Program Performance" in wb.sheets else wb.sheets[0]
        with wb.get_sheet(sheet_name) as s:
            rows = []
            for i, row in enumerate(s.rows()):
                r_vals = [c.v for c in row]
                # convert Excel serial dates
                if i == 0:
                    for idx, val in enumerate(r_vals):
                        if isinstance(val, str) and "data updated" in val.lower():
                            for k in range(idx + 1, min(idx + 4, len(r_vals))):
                                if isinstance(r_vals[k], (int, float)) and r_vals[k] > 20000:
                                    r_vals[k] = (datetime.date(1899, 12, 30) + datetime.timedelta(days=int(r_vals[k]))).strftime("%Y-%m-%d")
                                    break
                if i >= 2 and len(r_vals) > 0 and isinstance(r_vals[0], (int, float)) and r_vals[0] > 20000:
                    r_vals[0] = (datetime.date(1899, 12, 30) + datetime.timedelta(days=int(r_vals[0]))).strftime("%Y-%m-%d")
                rows.append(r_vals)
            embed_sample_data("program_performance.html", rows)

def sync_branch(opener):
    print("\n--- Syncing branch.html (Branch Performance) ---")
    data = fetch_url_data(opener, SHAREPOINT_URLS["branch"], LOCAL_FALLBACKS.get("branch"), label="branch")
    if not data:
        print("Warning: Could not fetch Branch Performance data.")
        return

    with open_workbook(io.BytesIO(data)) as wb:
        with wb.get_sheet("Branch Performance") as sheet:
            rows = []
            for i, row in enumerate(sheet):
                r_vals = [cell.v for cell in row]
                # Format serial date in row 0 if present
                if i == 0:
                    for idx, val in enumerate(r_vals):
                        if isinstance(val, str) and "data updated" in val.lower():
                            for k in range(idx + 1, min(idx + 4, len(r_vals))):
                                if isinstance(r_vals[k], (int, float)) and r_vals[k] > 20000:
                                    r_vals[k] = (datetime.date(1899, 12, 30) + datetime.timedelta(days=int(r_vals[k]))).strftime("%Y-%m-%d")
                                    break
                    rows.append(r_vals)
                elif i == 1:
                    rows.append(r_vals)
                # For data rows, ensure branch name (column 6) exists and is not blank
                elif len(r_vals) > 6 and r_vals[6] and str(r_vals[6]).strip():
                    rows.append(r_vals)
            embed_sample_data("branch.html", rows)

def auto_push_to_github():
    print("\n--- Pushing updates to GitHub ---")
    try:
        subprocess.run(["git", "add", "index.html", "training.html", "m_score.html", "program_performance.html", "branch.html", "sync_datasets.py"], check=True)
        # Check if there are changes to commit
        res = subprocess.run(["git", "diff", "--staged", "--quiet"])
        if res.returncode != 0:
            msg = f"chore(data): auto-sync latest live SharePoint datasets ({datetime.datetime.now().strftime('%Y-%m-%d %H:%M')})"
            subprocess.run(["git", "commit", "-m", msg], check=True)
            # Rebase onto latest remote in case remote was updated by GitHub Actions
            subprocess.run(["git", "pull", "--rebase", "origin", "main"], check=True)
            subprocess.run(["git", "push", "origin", "main"], check=True)
            print("🚀 Successfully pushed updated datasets to GitHub main branch!")
        else:
            # Still attempt to push any previously committed local syncs if ahead
            ahead_check = subprocess.run(["git", "log", "origin/main..main", "--oneline"], capture_output=True, text=True)
            if ahead_check.stdout.strip():
                subprocess.run(["git", "pull", "--rebase", "origin", "main"], check=True)
                subprocess.run(["git", "push", "origin", "main"], check=True)
                print("🚀 Successfully pushed pending local commits to GitHub main branch!")
            else:
                print("No data changes detected. Remote repository is already up-to-date.")
    except Exception as e:
        print("Note: Could not push to git automatically:", e)

def run_sync_once():
    print(f"[{datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] Starting SharePoint / Local Data Sync...")
    # Ensure local repo is up-to-date with remote before starting sync
    try:
        subprocess.run(["git", "pull", "--rebase", "origin", "main"], check=True)
    except Exception as e:
        print("Notice: Could not pull latest changes from remote, continuing with local base:", e)

    opener = get_opener()
    sync_tasks = [
        ("index.html (QC Tracker)", sync_index),
        ("training.html (Training Details)", sync_training),
        ("m_score.html (Product-VM-Score)", sync_m_score),
        ("program_performance.html (VM Performance)", sync_program_performance),
        ("branch.html (Branch Performance)", sync_branch),
    ]

    success_count = 0
    for name, func in sync_tasks:
        try:
            func(opener)
            success_count += 1
        except Exception as err:
            print(f"❌ Error syncing {name}: {err}")

    if success_count > 0:
        print(f"\nAll datasets processed ({success_count}/{len(sync_tasks)} succeeded)!")
        auto_push_to_github()
    else:
        print("\nNo datasets were synchronized.")

def watch_mode(interval=5):
    print(f"👀 Watching local OneDrive folder for instant changes:")
    print(f"   {ONEDRIVE_BASE}")
    print("Whenever you save any report file in Excel, it will automatically sync and push to GitHub.")
    print("Press Ctrl+C to stop watching.\n")
    
    tracked_files = []
    for paths in LOCAL_FALLBACKS.values():
        for p in paths:
            if p.startswith(ONEDRIVE_BASE) and p not in tracked_files:
                tracked_files.append(p)

    last_mtimes = {}
    for p in tracked_files:
        if os.path.exists(p):
            last_mtimes[p] = os.path.getmtime(p)

    while True:
        try:
            time.sleep(interval)
            changed_file = None
            for p in tracked_files:
                if os.path.exists(p):
                    m = os.path.getmtime(p)
                    if p not in last_mtimes:
                        last_mtimes[p] = m
                        changed_file = p
                        break
                    elif m > last_mtimes[p]:
                        changed_file = p
                        last_mtimes[p] = m
                        break
            if changed_file:
                print(f"\n🔔 File change detected: {os.path.basename(changed_file)}")
                print("Debouncing 3 seconds to allow Excel to finish writing...")
                time.sleep(3)
                run_sync_once()
        except KeyboardInterrupt:
            print("\nExiting watcher.")
            break
        except Exception as e:
            print(f"Watcher exception: {e}")
            time.sleep(interval)

def main():
    if "--watch" in sys.argv:
        watch_mode()
    else:
        run_sync_once()

if __name__ == "__main__":
    main()

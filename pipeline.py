#!/usr/bin/env python3
"""
BC Court Services Online - Daily Civil Case Scraper & Pipeline
Fetches the daily new case report PDF, extracts structured case entries,
updates the SQLite database with rolling retention, and exports spreadsheets
(Excel + CSV) and JSON index for search dashboard.
"""

import os
import re
import sys
import json
import sqlite3
import logging
from datetime import datetime, timezone
import requests
import pdfplumber
import pandas as pd

logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s")

BASE_URL = "https://justice.gov.bc.ca/cso"
INDEX_URL = f"{BASE_URL}/index.do"
REPORT_PAGE_URL = f"{BASE_URL}/newCaseReport.do"
PDF_URL = f"{BASE_URL}/viewNewCaseReport.do"

DATA_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "data")
DB_PATH = os.path.join(DATA_DIR, "cases.db")
EXCEL_PATH = os.path.join(DATA_DIR, "bc_civil_cases.xlsx")
CSV_PATH = os.path.join(DATA_DIR, "bc_civil_cases.csv")
JSON_PATH = os.path.join(DATA_DIR, "cases_search_index.json")

# Configurable rolling retention limit to keep DB and dashboard lightweight
# Keeps the latest MAX_RECORDS in the active search index / live dashboard
MAX_LIVE_RECORDS = int(os.environ.get("MAX_LIVE_RECORDS", "15000"))


def get_http_session():
    session = requests.Session()
    session.headers.update({
        "User-Agent": (
            "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 "
            "(KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
        ),
        "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8",
        "Accept-Language": "en-US,en;q=0.9",
    })
    return session


def download_pdf(output_pdf_path):
    logging.info("Initiating CSO session to acquire authentication cookies...")
    session = get_http_session()

    # Step 1: visit index
    resp1 = session.get(INDEX_URL, timeout=30)
    resp1.raise_for_status()

    # Step 2: visit report navigation page
    resp2 = session.get(REPORT_PAGE_URL, timeout=30)
    resp2.raise_for_status()

    # Step 3: fetch actual PDF report
    logging.info("Downloading daily PDF report from %s ...", PDF_URL)
    resp3 = session.get(PDF_URL, headers={"Referer": REPORT_PAGE_URL}, timeout=60)
    resp3.raise_for_status()

    content = resp3.content
    if not content.startswith(b"%PDF"):
        raise ValueError(f"Downloaded content is not a valid PDF! Response begins with: {content[:100]!r}")

    with open(output_pdf_path, "wb") as f:
        f.write(content)

    logging.info("Downloaded %d bytes to %s", len(content), output_pdf_path)
    return output_pdf_path


def parse_date(date_str):
    # Formats e.g. 21-SEP-2026 -> 2026-09-21
    try:
        dt = datetime.strptime(date_str.strip().upper(), "%d-%b-%Y")
        return dt.strftime("%Y-%m-%d")
    except Exception:
        return date_str.strip()


def parse_pdf(pdf_path):
    logging.info("Parsing PDF: %s", pdf_path)
    records = []
    current_court = ""

    with pdfplumber.open(pdf_path) as pdf:
        total_pages = len(pdf.pages)
        logging.info("PDF has %d pages", total_pages)

        for p_idx, page in enumerate(pdf.pages):
            # Crop off header (y <= 110) and footer (y >= 705)
            cropped = page.crop((0, 110, page.width, 705))
            words = cropped.extract_words(x_tolerance=2)

            # Dates opened serve as vertical row anchors
            date_words = [
                w for w in words
                if re.match(r"^\d{2}-[A-Z]{3}-\d{4}$", w["text"]) and w["x0"] > 500
            ]

            for i, dw in enumerate(date_words):
                row_top = dw["top"] - 4
                row_bottom = date_words[i + 1]["top"] - 4 if i + 1 < len(date_words) else 705

                row_words = [w for w in words if row_top <= w["top"] < row_bottom]

                # Court location in column 0 (x0 < 125)
                court_items = [w["text"] for w in row_words if w["x0"] < 125]
                fileno_items = [w["text"] for w in row_words if 125 <= w["x0"] < 195]

                # Handle concatenated digits if court label overflowed into file number column
                cleaned_court = []
                for c in court_items:
                    m = re.match(r"^(.*?)(\d+)$", c)
                    if m:
                        if m.group(1):
                            cleaned_court.append(m.group(1))
                        if not fileno_items:
                            fileno_items.append(m.group(2))
                    else:
                        cleaned_court.append(c)

                if cleaned_court:
                    court_candidate = " ".join(cleaned_court).strip()
                    if court_candidate:
                        current_court = court_candidate

                class_items = [w["text"] for w in row_words if 195 <= w["x0"] < 308]
                style_items = [w["text"] for w in row_words if 308 <= w["x0"] < 458]
                elec_items = [w["text"] for w in row_words if 458 <= w["x0"] < 500]
                date_items = [w["text"] for w in row_words if 500 <= w["x0"]]

                raw_date = " ".join(date_items).strip()
                norm_date = parse_date(raw_date)

                fileno = " ".join(fileno_items).strip()
                classification = " ".join(class_items).strip()
                style_of_cause = " ".join(style_items).strip()
                electronic_docs = " ".join(elec_items).strip() or "YES"

                if fileno and style_of_cause:
                    records.append({
                        "court_location": current_court,
                        "file_number": fileno,
                        "classification": classification,
                        "style_of_cause": style_of_cause,
                        "electronic_docs": electronic_docs,
                        "date_opened": norm_date,
                        "date_opened_raw": raw_date,
                    })

    logging.info("Extracted %d valid case rows from PDF", len(records))
    return records


def init_db(db_path):
    os.makedirs(os.path.dirname(db_path), exist_ok=True)
    conn = sqlite3.connect(db_path)
    cur = conn.cursor()
    cur.execute("""
        CREATE TABLE IF NOT EXISTS cases (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            court_location TEXT NOT NULL,
            file_number TEXT NOT NULL,
            classification TEXT,
            style_of_cause TEXT,
            electronic_docs TEXT,
            date_opened TEXT,
            date_opened_raw TEXT,
            scraped_at TEXT,
            UNIQUE(court_location, file_number, date_opened)
        )
    """)
    cur.execute("CREATE INDEX IF NOT EXISTS idx_cases_date ON cases(date_opened DESC)")
    cur.execute("CREATE INDEX IF NOT EXISTS idx_cases_court ON cases(court_location)")
    cur.execute("CREATE INDEX IF NOT EXISTS idx_cases_classification ON cases(classification)")
    cur.execute("CREATE INDEX IF NOT EXISTS idx_cases_fileno ON cases(file_number)")
    conn.commit()
    return conn


def save_to_db(records, db_path=DB_PATH):
    conn = init_db(db_path)
    cur = conn.cursor()
    now_iso = datetime.now(timezone.utc).isoformat()

    inserted = 0
    updated = 0
    for r in records:
        cur.execute("""
            INSERT INTO cases (
                court_location, file_number, classification,
                style_of_cause, electronic_docs, date_opened,
                date_opened_raw, scraped_at
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?)
            ON CONFLICT(court_location, file_number, date_opened) DO UPDATE SET
                classification = excluded.classification,
                style_of_cause = excluded.style_of_cause,
                electronic_docs = excluded.electronic_docs,
                date_opened_raw = excluded.date_opened_raw
        """, (
            r["court_location"],
            r["file_number"],
            r["classification"],
            r["style_of_cause"],
            r["electronic_docs"],
            r["date_opened"],
            r["date_opened_raw"],
            now_iso
        ))
        if cur.rowcount == 1:
            inserted += 1
        else:
            updated += 1

    conn.commit()
    logging.info("Database updated: %d new inserted, %d existing updated", inserted, updated)

    # Database pruning & size control
    # Keeps the database lean, fast, and below GitHub repository size limits
    cur.execute("SELECT COUNT(*) FROM cases")
    total_count = cur.fetchone()[0]
    logging.info("Total cases currently in database: %d", total_count)

    # Apply rolling retention limit if configured
    if MAX_LIVE_RECORDS > 0 and total_count > (MAX_LIVE_RECORDS * 1.2):
        logging.info("Pruning database beyond retention limit (%d records)...", MAX_LIVE_RECORDS)
        cur.execute("""
            DELETE FROM cases WHERE id NOT IN (
                SELECT id FROM cases ORDER BY date_opened DESC, id DESC LIMIT ?
            )
        """, (MAX_LIVE_RECORDS,))
        conn.commit()
        cur.execute("VACUUM")
        conn.commit()
        cur.execute("SELECT COUNT(*) FROM cases")
        pruned_count = cur.fetchone()[0]
        logging.info("Database pruned to %d records", pruned_count)

    conn.close()


def export_data(db_path=DB_PATH):
    logging.info("Exporting cases to spreadsheet formats (Excel, CSV) and Dashboard JSON...")
    conn = sqlite3.connect(db_path)
    query = """
        SELECT
            court_location AS "Court Location",
            file_number AS "File Number",
            classification AS "Classification",
            style_of_cause AS "Style of Cause",
            electronic_docs AS "Electronic Docs",
            date_opened AS "Date Opened",
            scraped_at AS "Scraped At"
        FROM cases
        ORDER BY date_opened DESC, court_location ASC, file_number ASC
    """
    df = pd.read_sql_query(query, conn)
    conn.close()

    # Save to CSV
    df.to_csv(CSV_PATH, index=False, encoding="utf-8")
    logging.info("Exported CSV to %s (%d rows)", CSV_PATH, len(df))

    # Save to Excel with formatted headers and auto column widths
    with pd.ExcelWriter(EXCEL_PATH, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="New Cases", index=False)
        worksheet = writer.sheets["New Cases"]
        for col in worksheet.columns:
            max_len = max(len(str(cell.value or "")) for cell in col)
            col_letter = col[0].column_letter
            worksheet.column_dimensions[col_letter].width = min(max(max_len + 3, 12), 65)

    logging.info("Exported Excel spreadsheet to %s", EXCEL_PATH)

    # Save lightweight JSON search index for fast static / client-side search dashboard
    # Limit dashboard search payload to the most recent MAX_LIVE_RECORDS
    search_df = df.head(MAX_LIVE_RECORDS)
    records = search_df.to_dict(orient="records")

    metadata = {
        "updated_at": datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M UTC"),
        "total_records": len(df),
        "displayed_records": len(records),
        "court_locations": sorted(list(df["Court Location"].dropna().unique())),
        "classifications": sorted(list(df["Classification"].dropna().unique())),
        "cases": records
    }

    with open(JSON_PATH, "w", encoding="utf-8") as f:
        json.dump(metadata, f, separators=(",", ":"))

    logging.info("Exported search JSON index to %s (%d records)", JSON_PATH, len(records))


def run_pipeline(custom_pdf=None):
    os.makedirs(DATA_DIR, exist_ok=True)
    pdf_path = custom_pdf or os.path.join(DATA_DIR, "latest_report.pdf")

    if not custom_pdf:
        download_pdf(pdf_path)

    records = parse_pdf(pdf_path)
    if not records:
        logging.warning("No records extracted from PDF!")
        return

    save_to_db(records)
    export_data()
    logging.info("Pipeline completed successfully.")


if __name__ == "__main__":
    pdf_arg = sys.argv[1] if len(sys.argv) > 1 else None
    run_pipeline(pdf_arg)

# B.C. Civil Court New Case Report Automation & Search Dashboard

Automated scraper and interactive web search dashboard for British Columbia Court Services Online (CSO) Civil New Case Reports.

## Features

- **Automated Daily Extraction**: Runs via GitHub Actions Monday through Friday at 12:00 PM Pacific Time (19:00 UTC).
- **Session-Aware Downloader**: Authenticates with Court Services Online (`index.do` -> `newCaseReport.do` -> `viewNewCaseReport.do`) to acquire session cookies and download the daily PDF report.
- **Accurate PDF Column Parsing**: Parses coordinates and extracts:
  - `Court Location` (handles carryover registry headings and boundary digit separation)
  - `File Number`
  - `Classification` (Supreme Civil, Small Claims, Probate, Foreclosure, etc.)
  - `Style of Cause` (Plaintiff vs Defendant party names)
  - `Electronic Docs`
  - `Date Opened`
- **Spreadsheet Generation**:
  - Automatically exports formatted Excel ([`bc_civil_cases.xlsx`](data/bc_civil_cases.xlsx)) with adjusted column widths.
  - Generates CSV ([`bc_civil_cases.csv`](data/bc_civil_cases.csv)) for data analysis and quick import.
- **Database Size Management Solution**:
  - Incremental upsert deduplication on `(court_location, file_number, date_opened)`.
  - Configurable rolling retention limit (`MAX_LIVE_RECORDS`, default 15,000 cases).
  - Automated `VACUUM` and pruned search index so the frontend loads instantly with low memory overhead.
- **Interactive Search Dashboard**:
  - Static HTML5/Tailwind/JavaScript dashboard ([`index.html`](index.html)) ready to deploy for free on GitHub Pages.
  - Real-time text search across Style of Cause parties and file numbers.
  - Dropdown filters for Court Location registry and Case Classification.
  - Column sorting and pagination.
  - Direct download links for Excel and CSV files.

## Project Structure

```text
├── .github/workflows/
│   └── daily_report.yml       # GitHub Actions schedule & GitHub Pages deployment
├── data/
│   ├── cases.db               # SQLite database with indexed case records
│   ├── bc_civil_cases.xlsx    # Formatted Excel spreadsheet
│   ├── bc_civil_cases.csv     # CSV spreadsheet
│   └── cases_search_index.json# Search index for the web dashboard
├── index.html                 # Web dashboard interface
├── pipeline.py                # Main scraping, parsing, DB management, & export script
├── requirements.txt           # Python dependencies
└── README.md
```

## Running Locally

1. Create and activate a Python virtual environment:
   ```bash
   python3 -m venv .venv
   source .venv/bin/activate
   pip install -r requirements.txt
   ```

2. Run the pipeline:
   ```bash
   # Download today's live report from CSO and build data/spreadsheets
   python pipeline.py

   # Or parse a local PDF file
   python pipeline.py /path/to/report.pdf
   ```

3. Launch the web dashboard:
   ```bash
   python3 -m http.server 8000
   ```
   Open `http://localhost:8000` in your web browser.

## Deploying to GitHub Pages

1. Push this repository to GitHub.
2. Go to **Settings > Pages** in your GitHub repository.
3. Under **Build and deployment > Source**, select **GitHub Actions**.
4. The workflow in `.github/workflows/daily_report.yml` will automatically run every business day and deploy updates to your GitHub Pages site.

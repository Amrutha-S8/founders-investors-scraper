# Indian Startup Investment Scraper & Analyzer

A Python data pipeline that cleans, analyzes, and visualizes **3,043 real Indian startup investment deals** (2015–2020) into a professional multi-sheet Excel workbook.

## What It Does

- Reads raw startup funding CSV data
- Cleans messy data (removes URLs, fixes amounts, normalizes categories)
- Outputs a fully formatted Excel workbook with 4 analysis sheets

## Output: `startups_cleaned.xlsx`

| Sheet | Contents |
|---|---|
| **Dashboard** | KPI summary cards + Top 10 Cities + Top 10 Industries |
| **All Deals** | All 3,043 records with filters, color rows, formatted amounts |
| **Investor Analysis** | Top 50 investors ranked by deal count with totals |
| **Investment Stages** | Seed vs Series A/B/C breakdown with amounts |

## Key Stats from the Data

- **3,043** investment deals
- **$38B+** total capital invested
- **112** cities across India
- **Top cities**: Bangalore, Mumbai, New Delhi, Gurgaon, Pune

## How to Run

### 1. Clone the repo
```bash
git clone https://github.com/Amrutha-S8/founders-investors-scraper.git
cd founders-investors-scraper
```

### 2. Install dependencies
```bash
pip install -r requirements.txt
```

### 3. Run the script
```bash
python scraper.py
```

### 4. Open the output
Open `startups_cleaned.xlsx` in Microsoft Excel or Google Sheets.

## Requirements

```
pandas
openpyxl
```

## Data Source

Indian startup funding data from Kaggle's Indian Startup Funding dataset, covering investments from 2015 to 2020.

## Files

```
founders-investors-scraper/
├── scraper.py              ← Main Python script
├── database.csv            ← Raw input data (3,043 rows)
├── startups_cleaned.xlsx   ← Formatted Excel output
├── requirements.txt        ← Python dependencies
└── README.md               ← This file
```

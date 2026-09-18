# EDGAR Financial Statement Scraper

Pulls balance sheet, income statement, and cash flow data for any public
company straight from SEC EDGAR, and saves it as a formatted Excel file.

## Setup

1. **Clone or download this repo.**

2. **Install dependencies:**
```bash
   pip install -r requirements.txt
```

3. **Add your info to `my_headers.py`:**
   SEC requires every request to identify who's making it. Open
   `my_headers.py` and replace the placeholder with your name and email:
```python
   headers = {'User-Agent': 'Your Name your.email@example.com'}
```

4. **Run it:**
```bash
   python edgar_functions.py
```
   You'll be prompted for a ticker and whether you want annual (10-K)
   or quarterly (10-Q) data. The output Excel file is saved to a
   `Company_Financials` folder on your Desktop.

## Optional: Mac keyboard shortcut

`scraper.command` (macOS only) lets you trigger the scraper from
anywhere via [Raycast](https://raycast.com), without opening a terminal
manually. Not required — the script above works fine on its own on
any OS. See [SETUP-RAYCAST.md] for how to wire it up.

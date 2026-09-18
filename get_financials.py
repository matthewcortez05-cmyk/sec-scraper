import pandas as pd
import numpy as np
import requests
import os
from bs4 import BeautifulSoup
import logging
import calendar
from my_headers import headers
from openpyxl.styles import Border, Side

pd.options.display.float_format = (
    lambda x: "{:,.0f}".format(x) if int(x) == x else "{:,.2f}".format(x)
)

statement_keys_map = {
    'balance_sheet': [
        "balance sheet",
        "balance sheets",
        "statement of financial position",
        "consolidated balance sheets",
        "consolidated financial position",
        "consolidated balance sheets - southern",
        "consolidated statements of financial position",
        "consolidated statement of financial position",
        "consolidated statements of financial condition",
        "combined and consolidated balance sheet",
        "condensed consolidated balance sheets",
        "consolidated balance sheets, as of december 31",
        "dow consolidated balance sheets",
        "consolidated balance sheets (unaudited)",
    ],
    "income_statement": [
        "income statement",
        "income statements",
        "statement of earnings (loss)",
        "statements of consolidated income",
        "consolidated statements of operations",
        "consolidated statement of operations",
        "consolidated statements of earnings",
        "consolidated statement of earnings",
        "consolidated statements of income",
        "consolidated statement of income",
        "consolidated income statements",
        "consolidated income statement",
        "condensed consolidated statements of earnings",
        "consolidated results of operations",
        "consolidated statements of income (loss)",
        "consolidated statements of income - southern",
        "consolidated statements of operations and comprehensive income",
        "consolidated statements of comprehensive income",
    ],
    "cash_flow_statement": [
        "cash flows statement",
        "cash flows statements",
        "statement of cash flows",
        "statements of consolidated cash flows",
        "consolidated statements of cash flows",
        "consolidated statement of cash flows",
        "consolidated statement of cash flow",
        "consolidated cash flows statements",
        "consolidated cash flow statements",
        "condensed consolidated statements of cash flows",
        "consolidated statements of cash flows (unaudited)",
        "consolidated statements of cash flows - southern",
    ],
}


def cik_matching_ticker(ticker, headers=headers):
    ticker = ticker.upper().replace("-",".")
    ticker_json = requests.get("https://www.sec.gov/files/company_tickers.json", headers=headers).json()

    for company in ticker_json.values():
        if company["ticker"] == ticker:
            cik = str(company["cik_str"]).zfill(10)
            return cik
        
        
    raise ValueError(f"Ticker {ticker} not found in SEC database")
    


def get_submission_data_for_ticker(ticker, headers=headers, only_filings_df=False):
    cik = cik_matching_ticker(ticker)
    url = f"https://data.sec.gov/submissions/CIK{cik}.json"
    company_json = requests.get(url, headers=headers).json()
    if only_filings_df:
        return pd.DataFrame(company_json['filings']['recent'])
    return company_json


def get_filtered_filings(ticker, ten_k=True, just_accession_numbers=False, headers=headers):
    company_filings_df = get_submission_data_for_ticker(ticker, only_filings_df=True, headers=headers)
    if ten_k:
        df = company_filings_df[company_filings_df['form'] == '10-K']
    else:
        df = company_filings_df[company_filings_df['form'] == '10-Q']
    if just_accession_numbers:
        df = df.set_index('reportDate')
        accession_df = df['accessionNumber']
        return accession_df
    else:
        return df
    

def get_facts(ticker, headers=headers):
    cik = cik_matching_ticker(ticker)
    url = f"https://data.sec.gov/api/xbrl/companyfacts/CIK{cik}.json"
    company_facts = requests.get(url, headers=headers).json()
    return company_facts


def facts_DF(ticker, headers=headers):
    facts = get_facts(ticker, headers)
    us_gaap_data = facts["facts"]["us-gaap"]
    df_data = []
    for fact, details in us_gaap_data.items():
        for unit in details["units"]:
            for item in details["units"][unit]:
                row = item.copy()
                row["fact"] = fact
                df_data.append(row)

    df = pd.DataFrame(df_data)
    df["end"] = pd.to_datetime(df["end"])
    df["start"] = pd.to_datetime(df["start"])
    df = df.drop_duplicates(subset=["fact", "end", "val"])
    labels_dict = {fact: details["label"] for fact, details in us_gaap_data.items()}
    return df, labels_dict


def annual_facts(ticker, headers=headers):
    accession_nums = get_filtered_filings(
        ticker, ten_k=True, just_accession_numbers=True
    )
    df, label_dict = facts_DF(ticker, headers)
    ten_k = df[df["accn"].isin(accession_nums)]
    pivot = ten_k.pivot_table(values="val", columns="fact", index="end")
    pivot.rename(columns=label_dict, inplace=True)
    return pivot.T


def quarterly_facts(ticker, headers=headers):
    accession_nums = get_filtered_filings(
        ticker, ten_k=False, just_accession_numbers=True
    )
    df, label_dict = facts_DF(ticker, headers)
    ten_q = df[df["accn"].isin(accession_nums)]
    ten_q = ten_q.drop_duplicates(subset=["fact", "end"], keep="last")
    pivot = ten_q.pivot_table(values="val", columns="fact", index="end")
    pivot.rename(columns=label_dict, inplace=True)
    return pivot.T


def save_dataframe_to_excel(dataframe, folder_name, ticker, statement_name, frequency):
    # Save to Desktop
    desktop_path = os.path.join(os.path.expanduser("~"), "Desktop")
    directory_path = os.path.join(desktop_path, folder_name, ticker)
    
    os.makedirs(directory_path, exist_ok=True)
    
    # Change extension to .xlsx
    file_path = os.path.join(directory_path, f"{statement_name}_{frequency}.xlsx")
    
    # Use to_excel instead of to_csv
    dataframe.to_excel(file_path)
    print(f"   Saved Excel: {file_path}")
    return None



def _get_file_name(report):
    html_file_name_tag = report.find("HtmlFileName")
    xml_file_name_tag = report.find("XmlFileName")

    if html_file_name_tag:
        return html_file_name_tag.text
    elif xml_file_name_tag:
        return xml_file_name_tag.text
    else:
        return ""


def _is_statement_file(short_name_tag, long_name_tag, file_name):
    return (
        short_name_tag is not None
        and long_name_tag is not None
        and file_name # Check if file_name is not an empty string
        and "Statement" in long_name_tag.text
    )




def get_statement_file_names_in_filing_summary(
    ticker, accession_number, headers=headers
):
    try:
        session = requests.Session()
        cik = cik_matching_ticker(ticker)
        base_link = f"https://www.sec.gov/Archives/edgar/data/{cik}/{accession_number}"
        filing_summary_link = f"{base_link}/FilingSummary.xml"
        filing_summary_response = session.get(
            filing_summary_link, headers=headers
        ).content.decode("utf-8")

        filing_summary_soup = BeautifulSoup(filing_summary_response, "lxml-xml")
        statement_file_names_dict = {}

        for report in filing_summary_soup.find_all("Report"):
            file_name = _get_file_name(report)
            short_name, long_name = report.find("ShortName"), report.find("LongName")

            if _is_statement_file(short_name, long_name, file_name):
                statement_file_names_dict[short_name.text.lower()] = file_name

        return statement_file_names_dict

    except requests.RequestException as e:
        print(f"An error occurred: {e}")
        return {}



def get_statement_soup(
    ticker,
    accession_number,
    statement_name,
    headers,
    statement_keys_map,
):
    """
    the statement_name should be one of the following:
    'balance_sheet'
    'income_statement'
    'cash_flow_statement'
    """
    session = requests.Session()

    cik = cik_matching_ticker(ticker)
    base_link = f"https://www.sec.gov/Archives/edgar/data/{cik}/{accession_number}"

    statement_file_name_dict = get_statement_file_names_in_filing_summary(
        ticker, accession_number, headers
    )

    statement_link = None
    for possible_key in statement_keys_map.get(statement_name.lower(), []):
        file_name = statement_file_name_dict.get(possible_key.lower())
        if file_name:
            statement_link = f"{base_link}/{file_name}"
            break

    if not statement_link:
        raise ValueError(f"Could not find statement file name for {statement_name}")

    try:
        statement_response = session.get(statement_link, headers=headers)
        statement_response.raise_for_status()  # Check if the request was successful

        if statement_link.endswith(".xml"):
            return BeautifulSoup(
                statement_response.content, "lxml-xml", from_encoding="utf-8"
            )
        else:
            return BeautifulSoup(statement_response.content, "lxml")

    except requests.RequestException as e:
        raise ValueError(f"Error fetching the statement: {e}")
    


def extract_columns_values_and_dates_from_statement(soup):
    columns = []
    values_set = []
    date_time_index = get_datetime_index_dates_from_statement(soup)

    for table in soup.find_all("table"):
        
        # We will default to 1, aiming to keep everything in Millions
        unit_multiplier = 1
        special_case = False

        # Grab the top text of the table
        table_text = table.get_text(separator=" ").lower()[:1000]
        
        if "in millions" in table_text:
            unit_multiplier = 1       # Leave millions exactly as they are reported
        elif "in thousands" in table_text:
            unit_multiplier = 0.001   # Shrink thousands down to millions so they match!
            
        if "unless otherwise specified" in table_text:
            special_case = True

        for row in table.select("tr"):
            onclick_elements = row.select("td.pl a, td.pl.custom a")
            if not onclick_elements:
                continue

            onclick_attr = onclick_elements[0]["onclick"]
            column_title = onclick_attr.split("defref_")[-1].split("',")[0]
            columns.append(column_title)

            # --- NEW: Identify Per-Share rows so we never scale EPS! ---
            is_per_share = "pershare" in column_title.lower() or "dividend" in column_title.lower()

            values = [np.nan] * len(date_time_index)

            for i, cell in enumerate(row.select("td.text, td.nump, td.num")):
                if "text" in cell.get("class"):
                    continue

                value = keep_numbers_and_decimals_only_in_string(
                    cell.text.replace("$", "")
                    .replace(",", "")
                    .replace("(", "")
                    .replace(")", "")
                    .strip()
                )
                
                if value:
                    value = float(value)
                    
                    # Apply multiplier ONLY if it's a standard account balance
                    actual_multiplier = 1 if is_per_share else unit_multiplier
                    
                    if special_case:
                        values[i] = value / 1000  
                    else:
                        if "nump" in cell.get("class"):
                            values[i] = value * actual_multiplier
                        else:
                            values[i] = -value * actual_multiplier

            values_set.append(values)

    return columns, values_set, date_time_index



def get_datetime_index_dates_from_statement(soup: BeautifulSoup) -> pd.DatetimeIndex:
    """
    Extracts datetime index dates from the HTML soup object of a financial statement.

    Args:
        soup (BeautifulSoup): The BeautifulSoup object of the HTML document.

    Returns:
        pd.DatetimeIndex: A Pandas DatetimeIndex object containing the extracted dates.
    """
    table_headers = soup.find_all("th", {"class": "th"})
    dates = [str(th.div.string) for th in table_headers if th.div and th.div.string]
    dates = [standardize_date(date).replace(".", "") for date in dates]
    index_dates = pd.to_datetime(dates)
    return index_dates


def standardize_date(date: str) -> str:
    """
    Standardizes date strings by replacing abbreviations with full month names.

    Args:
        date (str): The date string to be standardized.

    Returns:
        str: The standardized date string.
    """
    for abbr, full in zip(calendar.month_abbr[1:], calendar.month_name[1:]):
        date = date.replace(abbr, full)
    return date


def keep_numbers_and_decimals_only_in_string(mixed_string: str) -> str:
    """
    Filters a string to keep only numbers and decimal points.

    Args:
        mixed_string (str): The string containing mixed characters.

    Returns:
        str: String containing only numbers and decimal points.
    """
    num = "1234567890."
    allowed = list(filter(lambda x: x in num, mixed_string))
    return "".join(allowed)


def create_dataframe_of_statement_values_columns_dates(
    values_set, columns, index_dates
) -> pd.DataFrame:
    """
    Creates a DataFrame from statement values, columns, and index dates.

    Args:
        values_set (list): List of values for each column.
        columns (list): List of column names.
        index_dates (pd.DatetimeIndex): DatetimeIndex for the DataFrame index.

    Returns:
        pd.DataFrame: DataFrame constructed from the given data.
    """
    transposed_values_set = list(zip(*values_set))
    df = pd.DataFrame(transposed_values_set, columns=columns, index=index_dates)
    return df


def process_one_statement(ticker, accession_number, statement_name):
    """
    Processes a single financial statement identified by ticker, accession number, and statement name.

    Args:
        ticker (str): The stock ticker.
        accession_number (str): The SEC accession number.
        statement_name (str): Name of the financial statement.

    Returns:
        pd.DataFrame or None: DataFrame of the processed statement or None if an error occurs.
    """
    try:
        # Fetch the statement HTML soup
        soup = get_statement_soup(
            ticker,
            accession_number,
            statement_name,
            headers=headers,
            statement_keys_map=statement_keys_map,
        )
    except Exception as e:
        logging.error(
            f"Failed to get statement soup: {e} for accession number: {accession_number}"
        )
        return None

    if soup:
        try:
            # Extract data and create DataFrame
            columns, values, dates = extract_columns_values_and_dates_from_statement(
                soup
            )
            df = create_dataframe_of_statement_values_columns_dates(
                values, columns, dates
            )

            if not df.empty:
                # Remove duplicate columns
                df = df.T.drop_duplicates().T
                
                df = df.T
            else:
                logging.warning(
                    f"Empty DataFrame for accession number: {accession_number}"
                )
                return None
            
            return df
        except Exception as e:
            logging.error(f"Error processing statement: {e}")
            return None
        

def get_label_dictionary(ticker, headers):
    facts = get_facts(ticker, headers)
    us_gaap_data = facts["facts"]["us-gaap"]
    labels_dict = {fact: details["label"] for fact, details in us_gaap_data.items()}
    return labels_dict


def rename_statement(statement, label_dictionary):
    def get_best_label(raw_string):
        # 1. Catch any weird NaN or non-string values immediately
        if not isinstance(raw_string, str):
            return "Unknown_Account"
            
        # 2. Extract the core tag (e.g., turns "us-gaap_Revenues" into "Revenues")
        core_tag = raw_string.split("_", 1)[-1] if "_" in raw_string else raw_string
        
        # 3. Ask the dictionary for the translation
        translated_name = label_dictionary.get(core_tag)
        
        # 4. If the dictionary returns a valid name, use it!
        #    Otherwise, fall back to the core_tag so you NEVER get a blank row.
        if translated_name and str(translated_name).strip() != "":
            return translated_name
        else:
            # Adds spaces before capital letters so "AccountsPayable" becomes "Accounts Payable"
            import re
            clean_fallback = re.sub(r'(?<!^)(?=[A-Z])', ' ', core_tag)
            return clean_fallback

    # Apply our bulletproof logic to the index
    statement.index = statement.index.map(get_best_label)
    return statement




if __name__ == "__main__":
    
    # 1. Ask for Ticker and Timeframe
    ticker = input("Enter a stock ticker (e.g. AAPL): ").strip().upper()
    
    print("\nWhich timeframe do you want?")
    print("  [A] Annual (10-K)")
    print("  [Q] Quarterly (10-Q)")
    choice = input("Enter choice (A/Q): ").strip().upper()
    
    print(f"\n--- Processing Statements for {ticker} ---")
    
    try:
        # 2. Automatically fetch the latest Accession Number
        is_ten_k = (choice == 'A')
        
        print("Locating latest SEC filing...")
        accession_nums = get_filtered_filings(ticker, ten_k=is_ten_k, just_accession_numbers=True)
        
        if accession_nums.empty:
            print(f"❌ No filings found for {ticker}.")
            exit()
            
        # Grab the most recent accession number 
        latest_acc_num = accession_nums.iloc[0].replace("-", "") 
        report_date = accession_nums.index[0]
        report_type = "10-K" if is_ten_k else "10-Q"
        
        print(f"✅ Found latest {report_type} (Date: {report_date}) | Accession: {latest_acc_num}")

        # 3. Fetch the SEC label dictionary to translate the raw XBRL tags
        print("Fetching SEC label dictionary...")
        label_dict = get_label_dictionary(ticker, headers)

        # 4. Define the statements to pull
        statements_to_pull = {
            'income_statement': 'Income Statement',
            'balance_sheet': 'Balance Sheet',
            'cash_flow_statement': 'Cash Flow'
            }
        
        # 5. Scrape the data FIRST and store it in a dictionary
        successfully_scraped_data = {}
        
        for statement_key, sheet_title in statements_to_pull.items():
            print(f"Fetching {sheet_title}...")
            
            # Scrape the data
            df = process_one_statement(ticker, latest_acc_num, statement_key)
            
            if df is not None and not df.empty:
                # Rename the index (account names) using the dictionary
                df = rename_statement(df, label_dict)
                
                # Drop accounts (rows) that have no values for any dates
                df = df.dropna(how='all')
                
                
                # Format Date Columns based on report type
                if is_ten_k:
                    df.columns = pd.to_datetime(df.columns).strftime('%Y')
                else:
                    df.columns = pd.to_datetime(df.columns).strftime('%Y-%m-%d')
                
                # FLIP COLUMNS: Sorts dates oldest (left) to newest (right)
                df = df.sort_index(axis=1)
                
                # Save it to our holding dictionary if not entirely empty
                if not df.empty:
                    successfully_scraped_data[sheet_title] = df
                    print(f"  ✅ Successfully scraped: {sheet_title}")
                else:
                    print(f"  ⚠️ {sheet_title} was empty after dropping blank rows.")
            else:
                print(f"  ⚠️ Failed to scrape or parse {sheet_title}. (Data was empty or not found)")
        
        # 6. ONLY attempt to save the Excel file if we have at least one valid sheet
        if successfully_scraped_data:
            print("\nSaving to Excel...")
            desktop_path = os.path.join(os.path.expanduser("~"), "Desktop")
            directory_path = os.path.join(desktop_path, "Company_Financials", ticker)
            os.makedirs(directory_path, exist_ok=True)
            
            file_path = os.path.join(directory_path, f"{ticker}_{report_type}_{latest_acc_num}.xlsx")
            
            with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
                for sheet_name, dataframe in successfully_scraped_data.items():
                    # Write the dataframe to the sheet
                    dataframe.to_excel(writer, sheet_name=sheet_name)
                    
                    # Access the openpyxl worksheet object
                    worksheet = writer.sheets[sheet_name]
                    
                    # Define border styles
                    no_border = Border()
                    bottom_border = Border(bottom=Side(style='thin', color='000000'))

                    # 1. Reset all borders, keeping only the underline beneath the date header row
                    for row in worksheet.iter_rows():
                        for cell in row:
                            if cell.row == 1:
                                cell.border = bottom_border
                            else:
                                cell.border = no_border
                    
                    # 2. Autofit column widths and format numbers
                    for col in worksheet.columns:
                        max_length = 0
                        column_letter = col[0].column_letter # Gets 'A', 'B', 'C', etc.
                        
                        for cell in col:
                            # Number Formatting & Length Calculation
                            if isinstance(cell.value, (int, float)) and not pd.isna(cell.value):
                                if cell.value == int(cell.value):
                                    cell.number_format = '#,##0'
                                    cell_str_len = len(f"{cell.value:,.0f}") 
                                else:
                                    cell.number_format = '#,##0.00'
                                    cell_str_len = len(f"{cell.value:,.2f}")
                            else:
                                cell_str_len = len(str(cell.value)) if cell.value is not None else 0
                            
                            try:
                                if cell_str_len > max_length:
                                    max_length = cell_str_len
                            except Exception:
                                pass
                        
                        adjusted_width = (max_length + 2)
                        worksheet.column_dimensions[column_letter].width = adjusted_width
                        
            print(f"Success! Saved nicely formatted Excel file to:\n{file_path}")
            print(f"\nOpening Excel file...")

            import subprocess
            subprocess.run(['open', file_path])
            # ------------------------------------------------------------
            
        else:
            print("\n❌ Execution finished, but NO data was found. Excel file was not created.")

    except Exception as e:
        print(f"\n❌ Error during execution: {e}")
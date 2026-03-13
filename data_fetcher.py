# data_fetcher.py — Using yfinance for Indian stock financials

import yfinance as yf
import pandas as pd
import time

def fetch_company(company):
    """
    Fetches Income Statement, Balance Sheet, Cash Flow
    for one Indian company using yfinance
    """
    ticker = yf.Ticker(company["nse"])
    result = {}

    # ── Income Statement ──────────────────────
    try:
        inc = ticker.financials          # Annual income statement
        if inc is not None and not inc.empty:

            # yfinance gives columns as dates — convert to years
            inc.columns = [str(c.year) for c in inc.columns]

            # Pick the rows we care about
            rows_we_want = [
                "Total Revenue",
                "Gross Profit",
                "EBIT",
                "EBITDA",
                "Net Income",
                "Basic EPS",
            ]

            # Keep only rows that exist
            inc = inc[inc.index.isin(rows_we_want)]

            # Convert from Rupees to Crores (divide by 1 crore = 10,000,000)
            inc = inc / 1e7

            result["income"] = inc
            print(f"  ✅ Income statement fetched")
        else:
            print(f"  ⚠ No income statement available")
            result["income"] = pd.DataFrame()

    except Exception as e:
        print(f"  ✗ Income error: {e}")
        result["income"] = pd.DataFrame()

    # ── Balance Sheet ─────────────────────────
    try:
        bal = ticker.balance_sheet
        if bal is not None and not bal.empty:

            bal.columns = [str(c.year) for c in bal.columns]

            rows_we_want = [
                "Total Assets",
                "Total Debt",
                "Cash And Cash Equivalents",
                "Stockholders Equity",
                "Total Liabilities Net Minority Interest",
            ]

            bal = bal[bal.index.isin(rows_we_want)]
            bal = bal / 1e7

            result["balance"] = bal
            print(f"  ✅ Balance sheet fetched")
        else:
            print(f"  ⚠ No balance sheet available")
            result["balance"] = pd.DataFrame()

    except Exception as e:
        print(f"  ✗ Balance error: {e}")
        result["balance"] = pd.DataFrame()

    # ── Cash Flow ─────────────────────────────
    try:
        cf = ticker.cashflow
        if cf is not None and not cf.empty:

            cf.columns = [str(c.year) for c in cf.columns]

            rows_we_want = [
                "Operating Cash Flow",
                "Capital Expenditure",
                "Free Cash Flow",
                "Financing Cash Flow",
            ]

            cf = cf[cf.index.isin(rows_we_want)]
            cf = cf / 1e7

            result["cashflow"] = cf
            print(f"  ✅ Cash flow fetched")
        else:
            print(f"  ⚠ No cash flow available")
            result["cashflow"] = pd.DataFrame()

    except Exception as e:
        print(f"  ✗ Cashflow error: {e}")
        result["cashflow"] = pd.DataFrame()

    return result


def fetch_all(companies):
    """Fetches all companies one by one"""
    all_data = {}

    for co in companies:
        print(f"\n→ Fetching {co['name']}...")
        all_data[co["name"]] = fetch_company(co)
        time.sleep(2)    # Small pause between companies

    print("\n✓ All financial data fetched!")
    return all_data
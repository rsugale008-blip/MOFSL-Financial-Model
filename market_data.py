# market_data.py — Gets live stock prices

import yfinance as yf
import time

def fetch_market_data(companies):
    results = {}

    for co in companies:
        print(f"  Getting price for: {co['nse']}")
        try:
            ticker = yf.Ticker(co["nse"])
            info   = ticker.info

            results[co["name"]] = {
                "price":      info.get("currentPrice", 0),
                "mkt_cap_cr": round(info.get("marketCap", 0) / 1e7, 1),
                "beta":       info.get("beta", 1.0),
                "52w_high":   info.get("fiftyTwoWeekHigh", 0),
                "52w_low":    info.get("fiftyTwoWeekLow", 0),
                "pe_ratio":   info.get("trailingPE", 0),
                "pb_ratio":   info.get("priceToBook", 0),
            }
        except Exception as e:
            print(f"  ERROR: {e}")
            results[co["name"]] = {}

        time.sleep(1)

    return results
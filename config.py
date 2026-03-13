# config.py
TARGET = {
    "name":   "Motilal Oswal Financial Services",
    "av_symbol":  "MOTILALOFS.BSE",   # Alpha Vantage symbol
    "nse":        "MOTILALOFS.NS",    # yfinance symbol
    "nse_code":   "MOTILALOFS",       # nsetools symbol
}

COMPS = [
    {"name": "Nuvama Wealth Management", "av_symbol": "NUVAMA.BSE",     "nse": "NUVAMA.NS",      "nse_code": "NUVAMA"},
    {"name": "JM Financial",             "av_symbol": "JMFINANCIL.BSE", "nse": "JMFINANCIL.NS",  "nse_code": "JMFINANCIL"},
    {"name": "IIFL Capital Services",    "av_symbol": "IIFLCAPS.BSE",    "nse": "IIFLCAPS.NS",     "nse_code": "IIFLCAPS"},
    {"name": "Anand Rathi Wealth",       "av_symbol": "ANANDRATHI.BSE", "nse": "ANANDRATHI.NS",  "nse_code": "ANANDRATHI"},
]

ALL_COMPANIES = [TARGET] + COMPS

MARKET = {
    "risk_free_rate":   0.071,
    "equity_risk_prem": 0.075,
    "terminal_growth":  0.05,
    "tax_rate":         0.2517,
}
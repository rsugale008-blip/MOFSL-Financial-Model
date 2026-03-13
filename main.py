from config        import ALL_COMPANIES
from data_fetcher  import fetch_all
from market_data   import fetch_market_data
from model_builder import build_model

print("=" * 45)
print("   IB MODEL — PHASE 3: BUILDING MODEL")
print("=" * 45)

print("\n[1/3] Fetching financial data...")
raw_data = fetch_all(ALL_COMPANIES)

print("\n[2/3] Fetching market data...")
market_data = fetch_market_data(ALL_COMPANIES)

print("\n[3/3] Building Excel model...")
filename = build_model(raw_data, market_data)

print("\n" + "=" * 45)
print("  ✅ PHASE 3 COMPLETE!")
print(f"  Open this file: {filename}")
print("=" * 45)
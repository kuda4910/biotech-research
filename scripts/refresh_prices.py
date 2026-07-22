#!/usr/bin/env python3
"""
Refresh `price` (and per-position `updated`) fields in portfolio.json from live quotes.

Deliberately narrow in scope: this only ever touches the price/updated fields, never thesis
text, bear/mid/bull targets, catalysts, or open questions -- those are judgment calls that
belong in a reconciliation pass (see templates/News_Update_Prompt_Template.md), not a script.

Usage:
    python scripts/refresh_prices.py            # update portfolio.json in place
    python scripts/refresh_prices.py --dry-run  # print what would change, write nothing

No API key required -- uses Yahoo Finance's public chart endpoint with a browser-like
User-Agent (the endpoint 429s without one).
"""
import json
import sys
import urllib.request
import urllib.error
from datetime import date, timezone, datetime
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parent.parent
PORTFOLIO_PATH = REPO_ROOT / "portfolio.json"
QUOTE_URL = "https://query1.finance.yahoo.com/v8/finance/chart/{ticker}"
HEADERS = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"}


def fetch_price(ticker):
    req = urllib.request.Request(QUOTE_URL.format(ticker=ticker), headers=HEADERS)
    with urllib.request.urlopen(req, timeout=10) as resp:
        data = json.load(resp)
    result = data["chart"]["result"][0]
    meta = result["meta"]
    price = meta.get("regularMarketPrice")
    as_of = meta.get("regularMarketTime")
    as_of_str = (
        datetime.fromtimestamp(as_of, tz=timezone.utc).strftime("%Y-%m-%d %H:%M UTC")
        if as_of else "unknown"
    )
    return price, as_of_str


def main():
    dry_run = "--dry-run" in sys.argv

    portfolio = json.loads(PORTFOLIO_PATH.read_text(encoding="utf-8"))
    today = date.today().isoformat()
    any_changed = False

    for pos in portfolio["positions"]:
        ticker = pos["ticker"]
        try:
            price, as_of = fetch_price(ticker)
        except (urllib.error.URLError, KeyError, IndexError, ValueError) as e:
            print(f"  {ticker}: FAILED to fetch ({e}) -- left unchanged")
            continue

        if price is None:
            print(f"  {ticker}: no price in response -- left unchanged")
            continue

        old_price = pos.get("price")
        price = round(float(price), 2)
        changed = old_price != price
        marker = "->" if changed else "=="
        print(f"  {ticker}: {old_price} {marker} {price}  (quote as of {as_of})")

        if changed:
            any_changed = True
            pos["price"] = price
            pos["updated"] = today

    if any_changed:
        portfolio["updated"] = today

    if dry_run:
        print("\n--dry-run: no files written.")
        return

    if not any_changed:
        print("\nNo prices changed -- portfolio.json left untouched.")
        return

    PORTFOLIO_PATH.write_text(
        json.dumps(portfolio, indent=2, ensure_ascii=False) + "\n", encoding="utf-8"
    )
    print(f"\nWrote updates to {PORTFOLIO_PATH}")
    print("Review the diff (git diff portfolio.json), then commit/push yourself --")
    print("this script does not touch git.")


if __name__ == "__main__":
    main()

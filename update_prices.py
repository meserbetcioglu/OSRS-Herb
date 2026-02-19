#!/usr/bin/env python3
"""
Update both high and low prices from GE API.

Prices sheet structure:
  A: Name
  B: High Price
  C: Low Price
  D: Volume (1h)

Sheet1 formulas use config (T2/T3) to select which price column to reference.
"""

import requests
import xlwings as xw
from datetime import datetime
from pathlib import Path
import sys

def fetch_ge_data(discord_handle: str, email: str) -> tuple[dict, dict, dict]:
    """Fetch high prices, low prices, and volume from GE API."""
    headers = {
        'User-Agent': discord_handle,
        'From': email
    }
    # First fetch the mapping of item IDs to names
    print("  Fetching item mapping...")
    url_mapping = "https://prices.runescape.wiki/api/v1/osrs/mapping"
    response_mapping = requests.get(url_mapping, headers=headers)
    response_mapping.raise_for_status()
    mapping_data = response_mapping.json()
    
    id_to_name = {}
    for item in mapping_data:
        id_to_name[str(item["id"])] = item["name"]
    
    # Now fetch prices
    url = "https://prices.runescape.wiki/api/v1/osrs/latest"
    response = requests.get(url, headers=headers)
    response.raise_for_status()
    data = response.json()
    
    high_prices = {}
    low_prices = {}
    
    for item_id, item_data in data.get("data", {}).items():
        name = id_to_name.get(item_id)
        if name:
            high_prices[name] = item_data.get("high", 0)
            low_prices[name] = item_data.get("low", 0)
    
    # Fetch volume
    url_volume = "https://prices.runescape.wiki/api/v1/osrs/1h"
    response_vol = requests.get(url_volume, headers=headers)
    response_vol.raise_for_status()
    data_vol = response_vol.json()
    
    volume = {}
    for item_id, item_data in data_vol.get("data", {}).items():
        name = id_to_name.get(item_id)
        if name:
            # Sum high and low price volumes for total volume
            high_vol = item_data.get("highPriceVolume", 0)
            low_vol = item_data.get("lowPriceVolume", 0)
            volume[name] = high_vol + low_vol
    
    return high_prices, low_prices, volume, id_to_name.values()

def main():
    if len(sys.argv) > 1:
        excel_path = Path(sys.argv[1]).expanduser()
    else:
        base_dir = Path(__file__).parent
        excel_path = base_dir / "Herbology.xlsm"
        if not excel_path.exists():
            excel_path = base_dir / "Herbology.xlsx"
    
    # Open or connect to workbook
    book = None
    app = None
    app_was_created = False
    
    for app_iter in xw.apps:
        for book_iter in app_iter.books:
            if book_iter.fullname and book_iter.fullname.lower() == str(excel_path).lower():
                app = app_iter
                book = book_iter
                break
        if book:
            break

    if not book:
        app = xw.App(visible=False)
        book = app.books.open(str(excel_path))
        app_was_created = True

    ws_prices = book.sheets["Prices"]
    ws_config = book.sheets["Config"]
    
    # Read headers from Config sheet
    discord_handle = ws_config.range("B4").value
    email = ws_config.range("B5").value
    
    if not discord_handle or not email:
        print("ERROR: Discord handle and E-mail address must be provided in Config sheet!")
        print("  Discord handle (B4): {}".format(discord_handle if discord_handle else "MISSING"))
        print("  E-mail (B5): {}".format(email if email else "MISSING"))
        if app_was_created:
            book.close()
            app.quit()
        sys.exit(1)
    
    print(f"Using Discord: {discord_handle}")
    print(f"Using Email: {email}")
    
    # Fetch both high and low prices
    print("Fetching GE data...")
    high_prices, low_prices, volume, items = fetch_ge_data(discord_handle, email)
    
    print(f"  High prices: {len(high_prices)} items")
    print(f"  Low prices: {len(low_prices)} items")
    print(f"  Volume data: {len(volume)} items")
    
    # Update Prices sheet with both price types
    rows = []
    updated_count = 0
    for item_name in items:
        rows.append([
            item_name,
            high_prices.get(item_name, 0),
            low_prices.get(item_name, 0),
            volume.get(item_name, 0),
        ])
        if item_name in high_prices or item_name in low_prices:
            updated_count += 1

    if rows:
        ws_prices.range("A2").value = rows
    
    # Update timestamp in Prices sheet
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    ws_prices.range("G2").value = timestamp
    
    # Save
    book.save()
    
    # Only close/quit if we opened a new instance
    # If Excel was already open (via VBA), don't close the book - just release the reference
    if app_was_created:
        book.close()
        app.quit()
    
    print(f"\nUpdated {updated_count} items with both high/low prices")
    print(f"Last Updated: {timestamp}")
    print("Done!")

if __name__ == "__main__":
    main()

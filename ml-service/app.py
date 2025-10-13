# from flask import Flask, request, jsonify
# import os
# from dotenv import load_dotenv
# import google.generativeai as genai

# load_dotenv()

# app = Flask(__name__)

# API_KEY = os.environ.get("GEMINI_API_KEY")
# if not API_KEY:
#     raise RuntimeError("GEMINI_API_KEY not set. Run: export GEMINI_API_KEY=your_key_here")

# genai.configure(api_key=API_KEY)
# model = genai.GenerativeModel("models/gemini-2.5-pro")

# @app.route("/recommend", methods=["POST"])
# def recommend():
#     try:
#         data = request.get_json()
#         saved_posts = data.get("saved_posts", [])
#         properties = data.get("properties", [])

#         if not saved_posts or not properties:
#             return jsonify({"error": "Missing saved_posts or properties"}), 400

#         # Build Gemini prompt
#         prompt = "You are a recommendation engine for real estate listings.\n"
#         prompt += "The user has saved these posts:\n"

#         for sp in saved_posts:
#             pd = sp.get("postDetail", {})
#             prompt += (
#                 f"- ID {sp.get('id', 'N/A')}: {sp.get('title', 'No title')}, "
#                 f"Price: ${sp.get('price', 'N/A')}, Type: {sp.get('type', 'N/A')}, "
#                 f"Property: {sp.get('property', 'N/A')}, Bedrooms: {sp.get('bedroom', 'N/A')}, "
#                 f"Bathrooms: {sp.get('bathroom', 'N/A')}, Address: {sp.get('address', '')}, "
#                 f"City: {sp.get('city', '')}, Latitude: {sp.get('latitude', '')}, Longitude: {sp.get('longitude', '')}, "
#                 f"Category: {sp.get('category', 'N/A')}, Images: {', '.join(sp.get('images', []))}, "
#                 f"Description: {pd.get('desc', '')}, Utilities: {pd.get('utilities', '')}, "
#                 f"Pet: {pd.get('pet', '')}, Income: {pd.get('income', '')}, Size: {pd.get('size', '')}, "
#                 f"School: {pd.get('school', '')}, Bus: {pd.get('bus', '')}, Restaurant: {pd.get('restaurant', '')}\n"
#             )

#             prompt += "\nAvailable properties to recommend from:\n"
#             for p in properties:
#                 pd = p.get("postDetail", {})
#                 prompt += (
#                     f"- ID {p.get('id', 'N/A')}: {p.get('title', 'No title')}, "
#                     f"Price: ${p.get('price', 'N/A')}, Type: {p.get('type', 'N/A')}, "
#                     f"Property: {p.get('property', 'N/A')}, Bedrooms: {p.get('bedroom', 'N/A')}, "
#                     f"Bathrooms: {p.get('bathroom', 'N/A')}, Address: {p.get('address', '')}, "
#                     f"City: {p.get('city', '')}, Latitude: {p.get('latitude', '')}, Longitude: {p.get('longitude', '')}, "
#                     f"Category: {p.get('category', 'N/A')}, Images: {', '.join(p.get('images', []))}, "
#                     f"Description: {pd.get('desc', '')}, Utilities: {pd.get('utilities', '')}, "
#                     f"Pet: {pd.get('pet', '')}, Income: {pd.get('income', '')}, Size: {pd.get('size', '')}, "
#                     f"School: {pd.get('school', '')}, Bus: {pd.get('bus', '')}, Restaurant: {pd.get('restaurant', '')}\n"
#                 )

#             prompt += (
#                 "\nSuggest 3 property IDs that best match the user's interests. "
#                 "Return only a comma-separated list of IDs, no explanations."
#             )

#         # Ask Gemini
#         response = model.generate_content(prompt)
#         raw_text = (response.text or "").strip()

#         if not raw_text:
#             return jsonify({"error": "No recommendations generated"}), 500

#         # Parse output into list
#         recommended_ids = [pid.strip() for pid in raw_text.split(",") if pid.strip()]
#         recommendations = recommended_ids[:3]

#         return jsonify({"recommendations": recommendations})

#     except Exception as e:
#         print(f"Error in /recommend: {e}")
#         return jsonify({"error": str(e)}), 500

# if __name__ == "__main__":
#     app.run(host="0.0.0.0", port=5001, debug=True)



#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
scrape_asins.py

- Loads your Google Merchant feed (google_merchant_feed_own.tsv).
- For each product title, searches Amazon.in directly.
- Scrapes Top 5 ASINs from search results.
- Writes them into new columns competitor_asin_1..5.
- Saves to competition_scraped.xlsx

Dependencies:
    pip install pandas requests beautifulsoup4 lxml openpyxl
"""

import os
import time
import random
import pandas as pd
import requests
from bs4 import BeautifulSoup

# ---------------------
# Config
# ---------------------
FEED_TSV = "google_merchant_feed_own.tsv"
OUTPUT_EXCEL = "competition_scraped.xlsx"
TOP_N = 5
REQUEST_DELAY = (2, 5)  # random delay between 2–5 sec per request

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/114.0 Safari/537.36"
    )
}

# ---------------------
# Helpers
# ---------------------
def extract_asins_from_html(html, limit=5):
    """Parse Amazon search page HTML and return up to `limit` ASINs."""
    soup = BeautifulSoup(html, "html.parser")
    asins = []
    for div in soup.find_all("div", {"data-asin": True}):
        asin = div.get("data-asin")
        if asin and asin.strip() and asin not in asins:
            asins.append(asin.strip())
        if len(asins) >= limit:
            break
    return asins

def search_amazon_asins(query, top_n=5):
    """Perform an Amazon search and extract ASINs."""
    url = "https://www.amazon.in/s"
    params = {"k": query}
    try:
        resp = requests.get(url, params=params, headers=HEADERS, timeout=15)
        if resp.status_code != 200:
            print(f"❌ HTTP {resp.status_code} for query: {query}")
            return []
        return extract_asins_from_html(resp.text, limit=top_n)
    except requests.RequestException as e:
        print(f"❌ Request error for query '{query}': {e}")
        return []

# ---------------------
# Main
# ---------------------
def main():
    if not os.path.exists(FEED_TSV):
        print(f"❌ Missing {FEED_TSV}")
        return

    df = pd.read_csv(FEED_TSV, sep="\t", dtype=str)
    df.columns = [c.strip().lower() for c in df.columns]

    # Add competitor columns
    for i in range(1, TOP_N + 1):
        df[f"competitor_asin_{i}"] = ""

    # Process rows
    for idx, row in df.iterrows():
        title = (row.get("title") or "").strip()
        brand = (row.get("brand") or "").strip()
        if not title:
            continue

        query = f"{brand} {title}".strip()
        print(f"🔎 Searching Amazon for: {query}")
        asins = search_amazon_asins(query, top_n=TOP_N)

        for i, asin in enumerate(asins):
            df.at[idx, f"competitor_asin_{i+1}"] = asin

        print(f"  → Found {len(asins)} ASINs: {asins}")
        time.sleep(random.uniform(*REQUEST_DELAY))

    # Save
    df.to_excel(OUTPUT_EXCEL, index=False)
    print(f"✅ Done. Results saved to {OUTPUT_EXCEL}")

if __name__ == "__main__":
    main()

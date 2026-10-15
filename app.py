
import re
import sys
from pathlib import Path
from urllib.parse import quote_plus
import xml.etree.ElementTree as ET

import pandas as pd
import requests


NEWS_RSS_URL = "https://news.google.com/rss/search"
REQUEST_TIMEOUT = 15


def get_google_news_result(term, count):
    """Fetch article titles and links from Google News RSS."""
    params = {
        "q": term,
        "hl": "en-IN",
        "gl": "IN",
        "ceid": "IN:en",
    }

    try:
        response = requests.get(
            NEWS_RSS_URL,
            params=params,
            timeout=REQUEST_TIMEOUT,
            headers={"User-Agent": "GoogleNewsScraper/1.0"},
        )
        response.raise_for_status()

        root = ET.fromstring(response.content)

    except requests.RequestException as exc:
        print(f"Network error while fetching news: {exc}")
        return []

    except ET.ParseError:
        print("Could not read the news feed. Please try again.")
        return []

    articles = []
    seen_links = set()

    for item in root.findall("./channel/item"):
        title = item.findtext("title", default="").strip()
        link = item.findtext("link", default="").strip()

        if title and link and link not in seen_links:
            articles.append({"title": title, "link": link})
            seen_links.add(link)

        if len(articles) >= count:
            break

    return articles


def create_output_filename(term):
    """Create a filesystem-safe Excel filename."""
    safe_term = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", term)
    safe_term = safe_term.strip(" ._") or "news"
    return f"{safe_term[:80]}_news_scrapper.xlsx"


def main():
    term = input("Enter the news title keyword: ").strip()

    if not term:
        print("Please enter a search keyword.")
        return

    try:
        count = int(input("Enter the number of articles: ").strip())
        if count < 1:
            print("Article count must be at least 1.")
            return
    except ValueError:
        print("Please enter a valid whole number.")
        return

    print(f"\nSearching Google News for: {term}")

    articles = get_google_news_result(term, count)

    if not articles:
        print("No articles found, or the news feed could not be retrieved.")
        return

    output_file = create_output_filename(term)

    try:
        df = pd.DataFrame(articles, columns=["title", "link"])
        df.to_excel(output_file, index=False)

    except (OSError, PermissionError, ValueError) as exc:
        print(f"Could not save the Excel file: {exc}")
        return

    print(f"\nSuccessfully saved {len(articles)} articles.")
    print(f"Excel file: {Path(output_file).resolve()}")


if __name__ == "__main__":
    main()

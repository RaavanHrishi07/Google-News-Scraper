
import re
from pathlib import Path

import pandas as pd
import requests
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
import xml.etree.ElementTree as ET


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


def save_articles_to_excel(articles, output_file):
    """Save articles in a formatted Excel workbook."""
    df = pd.DataFrame(articles, columns=["title", "link"])
    df.columns = ["Title", "Article Link"]

    with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="News Articles")

        worksheet = writer.sheets["News Articles"]
        worksheet.freeze_panes = "A2"
        worksheet.auto_filter.ref = worksheet.dimensions

        header_fill = PatternFill(
            fill_type="solid",
            fgColor="1F4E78",
        )

        for cell in worksheet[1]:
            cell.font = Font(
                bold=True,
                color="FFFFFF",
                size=12,
            )
            cell.fill = header_fill
            cell.alignment = Alignment(
                horizontal="center",
                vertical="center",
            )

        worksheet.row_dimensions[1].height = 25
        worksheet.column_dimensions["A"].width = 65
        worksheet.column_dimensions["B"].width = 55

        for row in range(2, worksheet.max_row + 1):
            title_cell = worksheet.cell(row=row, column=1)
            link_cell = worksheet.cell(row=row, column=2)

            title_cell.alignment = Alignment(
                vertical="top",
                wrap_text=True,
            )

            link_cell.hyperlink = link_cell.value
            link_cell.style = "Hyperlink"
            link_cell.alignment = Alignment(vertical="top")

            worksheet.row_dimensions[row].height = 30


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
        save_articles_to_excel(articles, output_file)

    except (OSError, PermissionError, ValueError) as exc:
        print(f"Could not save the Excel file: {exc}")
        return

    print(f"\nSuccessfully saved {len(articles)} articles.")
    print(f"Excel file: {Path(output_file).resolve()}")


if __name__ == "__main__":
    main()

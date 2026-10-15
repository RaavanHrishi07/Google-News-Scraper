
import re
from pathlib import Path
from urllib.parse import urlparse
import xml.etree.ElementTree as ET

import pandas as pd
import requests
from openpyxl.styles import Font, PatternFill, Alignment


NEWS_RSS_URL = "https://news.google.com/rss/search"
REQUEST_TIMEOUT = 15


def get_google_news_result(term, count):
    """Fetch news titles, links, sources, and publication dates."""
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
        published = item.findtext("pubDate", default="").strip()

        # Google News commonly includes the source in <source>.
        source_element = item.find("source")
        source = (
            source_element.text.strip()
            if source_element is not None and source_element.text
            else ""
        )

        if not source and link:
            source = urlparse(link).netloc.removeprefix("www.")

        if title and link and link not in seen_links:
            articles.append(
                {
                    "title": title,
                    "source": source or "Unknown",
                    "published_date": published or "Not available",
                    "link": link,
                }
            )
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
    """Save news articles in a formatted Excel workbook."""
    df = pd.DataFrame(
        articles,
        columns=["title", "source", "published_date", "link"],
    )

    df.columns = [
        "Title",
        "News Source",
        "Publication Date",
        "Article Link",
    ]

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
            cell.font = Font(bold=True, color="FFFFFF", size=12)
            cell.fill = header_fill
            cell.alignment = Alignment(
                horizontal="center",
                vertical="center",
            )

        worksheet.row_dimensions[1].height = 25

        column_widths = {
            "A": 60,
            "B": 28,
            "C": 35,
            "D": 55,
        }

        for column, width in column_widths.items():
            worksheet.column_dimensions[column].width = width

        for row in range(2, worksheet.max_row + 1):
            worksheet.cell(row=row, column=1).alignment = Alignment(
                vertical="top",
                wrap_text=True,
            )

            for column in (2, 3):
                worksheet.cell(row=row, column=column).alignment = (
                    Alignment(vertical="top", wrap_text=True)
                )

            link_cell = worksheet.cell(row=row, column=4)
            link_cell.hyperlink = link_cell.value
            link_cell.style = "Hyperlink"

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

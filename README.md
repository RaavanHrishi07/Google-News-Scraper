# 📰 Google News Scraper

A Python-based automation tool that fetches news articles from Google News RSS and exports the results to professionally formatted Excel and CSV files.

Search for any topic, choose how many articles you want, and generate a report containing article titles, news sources, publication dates, and links.

## ✨ Features

- 🔎 Search Google News using any keyword.
- 📰 Retrieve article titles, sources, publication dates, and links.
- 📊 Export news data to Excel (`.xlsx`) and CSV (`.csv`).
- 📋 Generate an Excel Summary sheet with the search keyword, article count, and report generation time.
- 🎨 Format Excel headers and adjust column widths for readability.
- 🔗 Create clickable article links in Excel.
- 🔍 Support Excel filters and frozen header rows.
- 🛡️ Handle network errors, invalid input, and duplicate article links.
- 🗂️ Generate filenames based on the search keyword.

## 🛠️ Tech Stack

- **Language:** Python
- **Data handling:** pandas
- **HTTP requests:** requests
- **Excel export and formatting:** openpyxl
- **News feed:** Google News RSS

## 📁 Project Structure

```text
Google-News-Scraper/
├── app.py
├── requirements.txt
├── README.md
├── LICENSE
└── .gitignore
```

Generated Excel and CSV reports are stored locally and excluded from Git.

## 🚀 Getting Started

### 1. Clone the repository

```bash
git clone https://github.com/RaavanHrishi07/Google-News-Scraper.git
cd Google-News-Scraper
```

### 2. Create a virtual environment (recommended)

On Windows PowerShell:

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
```

### 3. Install dependencies

```powershell
python -m pip install -r requirements.txt
```

### 4. Run the scraper

```powershell
python app.py
```

Enter a search keyword and the number of articles when prompted.

Example:

```text
Enter the news title keyword: cricket
Enter the number of articles: 10
```

## 📄 Output Files

For the example keyword `cricket`, the application generates:

- **`cricket_news_scrapper.xlsx`** — an Excel workbook containing:
  - `Summary`: search keyword, total articles, and report generation timestamp.
  - `News Articles`: article title, news source, publication date, and clickable link.
- **`cricket_news_scrapper.csv`** — a CSV export containing the article data.

Output filenames are based on the search keyword.

## ⚙️ How It Works

1. Accepts a search keyword and requested article count.
2. Requests the Google News RSS feed.
3. Extracts available article titles, links, sources, and publication dates.
4. Removes duplicate article links.
5. Exports the collected data to Excel and CSV.
6. Applies formatting to the Excel workbook.

Some source or publication-date information may be unavailable in the RSS feed.

## ⚠️ Notes

- An active internet connection is required.
- The number of returned articles may be lower than requested if fewer valid results are available.
- Google News RSS availability and response formats may change.
- Use the tool responsibly and respect applicable website terms and usage policies.

## 👨‍💻 Author

**RaavanHrishi07**

GitHub: [@RaavanHrishi07](https://github.com/RaavanHrishi07)

## 📜 License

Distributed under the MIT License. See [`LICENSE`](LICENSE) for details.
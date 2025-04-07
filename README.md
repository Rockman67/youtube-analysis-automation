# YouTube Channel Analysis Automation

This project contains Python scripts designed to automate the process of analyzing YouTube channels using the YouTube Data API and Selenium.

---

## Project Overview

The scripts automate:
- Retrieving channel handles from YouTube.
- Collecting detailed statistics including total videos, shorts, views, likes, comments, emails, and author location.
- Storing intermediate results in Excel and SQLite to prevent data loss.

---

## Files and Scripts

- **`sch.py`**: Initial script for searching YouTube videos based on keywords and initial filtering.
- **`test2.py`**: Advanced script for detailed channel analysis and data collection.
- **`channels_data.db`**: SQLite database for storing processed video IDs.
- **`channel_info.xlsx` / `final_channels.xlsx`**: Excel files for intermediate and final results.

---

## Requirements

- Python 3.x
- Google API Key (YouTube Data API v3)

Install Python dependencies with:
```bash
pip install google-api-python-client selenium webdriver-manager openpyxl langid pandas
```

Ensure you have Google Chrome installed for Selenium automation.

---

## Configuration

Before running the scripts:

1. Replace the placeholder API key (`DEVELOPER_KEY`) in `sch.py` and `test2.py` with your own YouTube Data API key.

2. Adjust settings such as:
   - `XLSX_INPUT` and `XLSX_OUTPUT` filenames in `test2.py`.
   - Keywords, date range, and subscriber limits in `sch.py`.

---

## Running the Scripts

### Initial Analysis (sch.py)

Run the initial script to search and filter YouTube videos:

```bash
python sch.py
```

This will generate/update the Excel file (`channel_info.xlsx`) containing basic channel data.

### Detailed Analysis (test2.py)

After obtaining basic data, run the detailed script to gather extended statistics:

```bash
python test2.py
```

This script produces detailed analytics in `final_channels.xlsx`, including email addresses and channel locations.

---

## Troubleshooting

- If you encounter quota limits (`quotaExceeded`), pause execution and retry after 24 hours.
- Check log outputs for error descriptions.

---

## Contact Information

For further assistance, development, or custom automation solutions, feel free to contact me:

- Linkedin: https://www.linkedin.com/in/aleksandr-boltenkov-124438353/

---

## License

Feel free to modify and adapt this code according to your needs. Please credit the original author if you reuse the scripts.

---

Enjoy automating your YouTube channel analysis!


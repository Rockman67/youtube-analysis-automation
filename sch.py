import os
import datetime
import langid
import pandas as pd
import sqlite3
import time
import logging
import re
import traceback

from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

# For Selenium
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service as ChromeService
from webdriver_manager.chrome import ChromeDriverManager
from requests.exceptions import ConnectionError as RequestsConnectionError
import urllib3.exceptions

# ------------------------------------------------------------------------------
# FUNCTION: Retries for YouTube API (search.list and channels.list)
# ------------------------------------------------------------------------------
def youtube_api_call_with_retries(api_func, max_retries=3, sleep_seconds=5):
    """
    Calls api_func() (which should return an object on which .execute() is called),
    and performs multiple retries in case of HttpError/connection issues
    in order to bypass transient failures (ConnectionAbortedError, RemoteDisconnected, etc.).

    Returns the response or None if all max_retries fail.
    """
    for attempt in range(1, max_retries + 1):
        try:
            response = api_func().execute()
            return response
        except (HttpError, ConnectionAbortedError, OSError,
                urllib3.exceptions.ProtocolError,
                RequestsConnectionError) as e:
            logging.error(f"[youtube_api_call_with_retries] Attempt {attempt}/{max_retries} -> error: {e}")
            if attempt < max_retries:
                logging.info(f"Waiting {sleep_seconds} seconds and then will retry...")
                time.sleep(sleep_seconds)
            else:
                logging.error("Retry limit exceeded. Returning None.")
                return None
        except Exception as e2:
            logging.error(f"[youtube_api_call_with_retries] Unexpected error: {e2}")
            traceback.print_exc()
            return None
    return None


# ------------------------------------------------------------------------------
# FUNCTION: Get channel handle via Selenium (with retries)
# ------------------------------------------------------------------------------
def get_handle_from_channel_id_selenium(channel_id: str,
                                        max_retries=3,
                                        sleep_seconds=5) -> str:
    """
    Opens https://www.youtube.com/channel/<channel_id> in Selenium,
    looks for selector:
       div.yt-content-metadata-view-model-wiz__metadata-row
           .yt-content-metadata-view-model-wiz__metadata-row--metadata-row-inline
           span.yt-core-attributed-string--link-inherit-color

    Returns something like '@Evel-901' or None if not found.
    Tries up to max_retries times if webdriver_manager or the browser
    throw a network error.
    """

    for attempt in range(1, max_retries + 1):
        try:
            return _try_open_channel_and_get_handle(channel_id)
        except (RequestsConnectionError,
                urllib3.exceptions.ProtocolError,
                ConnectionAbortedError,
                Exception) as e:
            logging.error(f"[get_handle_from_channel_id_selenium] Attempt {attempt}/{max_retries} -> Error: {e}")
            traceback.print_exc()
            if attempt < max_retries:
                logging.info(f"Waiting {sleep_seconds} seconds and then will retry Selenium...")
                time.sleep(sleep_seconds)
            else:
                logging.error("Selenium retry limit exceeded.")
                return None
    return None


def _try_open_channel_and_get_handle(channel_id: str) -> str:
    """
    Actual logic for opening the browser + finding the handle.
    Called by get_handle_from_channel_id_selenium in a retry loop.
    """
    logging.info(f"[_try_open_channel_and_get_handle] Starting for channel_id={channel_id}")

    # --- ChromeOptions settings ---
    options = webdriver.ChromeOptions()
    # Enable headless mode if desired:
    # options.add_argument("--headless")

    # Increase timeouts:
    options.add_argument("--start-maximized")
    options.page_load_strategy = "normal"

    # If you already downloaded ChromeDriver manually, you can specify the path:
    # driver_path = r"C:\path\to\chromedriver.exe"
    # service = ChromeService(executable_path=driver_path)
    # driver = webdriver.Chrome(service=service, options=options)

    # Otherwise, try to download via webdriver_manager:
    service = ChromeService(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=options)

    # Set large timeouts
    driver.set_page_load_timeout(120)  # up to 120 seconds waiting for page load
    driver.implicitly_wait(30)         # up to 30 seconds waiting for elements
    logging.info("[Selenium] WebDriver started, opening page...")

    try:
        url = f"https://www.youtube.com/channel/{channel_id}"
        driver.get(url)

        # Additional pause to ensure everything is loaded (banner, JS, etc.)
        time.sleep(15)

        logging.info("[Selenium] Looking for the required <span>...")

        span_handle = driver.find_element(
            By.CSS_SELECTOR,
            "div.yt-content-metadata-view-model-wiz__metadata-row"
            ".yt-content-metadata-view-model-wiz__metadata-row--metadata-row-inline "
            "span.yt-core-attributed-string--link-inherit-color"
        )

        found_handle = span_handle.text.strip()
        if found_handle:
            logging.info(f"[Selenium] Found handle: {found_handle}")
            return found_handle
        else:
            logging.warning("[Selenium] Span element found but text is empty.")
            return None

    finally:
        logging.info("[Selenium] Closing the browser.")
        driver.quit()


# ------------------------------------------------------------------------------
# MAIN FUNCTION
# ------------------------------------------------------------------------------
def main():
    logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s")

    # Insert your own API key here
    DEVELOPER_KEY = "YOUR_API_KEY_HERE"
    youtube = build("youtube", "v3", developerKey=DEVELOPER_KEY)

    # Search settings: last year, region FR
    days_back = 365
    published_after = datetime.datetime.utcnow() - datetime.timedelta(days=days_back)
    published_after_str = published_after.isoformat("T") + "Z"

    FRENCH_QUERIES = [
        "comment",
        "pourquoi",
        "ça",
        "français",
        "j'ai",
        "mes amis",
        "vous avez",
        "peut-être",
        "alors",
        "truc",
        "c'est",
        "faut"
    ]

    logging.info("=== Settings ===")
    logging.info(f"Searching for videos published after {published_after_str}, regionCode='FR', order='date', maxResults=50.")
    logging.info(f"French keywords: {FRENCH_QUERIES}")

    # ----------------------------------------------------------------------
    # Prepare DB (SQLite): "channels_data.db" with table "processed_videos"
    # ----------------------------------------------------------------------
    db_path = "channels_data.db"
    conn = sqlite3.connect(db_path)
    cur = conn.cursor()

    cur.execute("""
        CREATE TABLE IF NOT EXISTS processed_videos (
            video_id TEXT PRIMARY KEY
        )
    """)
    conn.commit()

    # ----------------------------------------------------------------------
    # Prepare Excel (channel_info.xlsx)
    # ----------------------------------------------------------------------
    excel_path = "channel_info.xlsx"
    if os.path.exists(excel_path):
        df_channels = pd.read_excel(excel_path, engine="openpyxl")
    else:
        df_channels = pd.DataFrame(columns=["channel_handle", "subscribers"])

    # Check columns
    if "channel_handle" not in df_channels.columns:
        df_channels["channel_handle"] = ""
    if "subscribers" not in df_channels.columns:
        df_channels["subscribers"] = 0

    total_videos_fetched = 0
    total_new_channels = 0

    # ----------------------------------------------------------------------
    # Loop over keywords, navigate search() results by pageToken
    # ----------------------------------------------------------------------
    for query_str in FRENCH_QUERIES:
        logging.info(f"=== Starting search for query '{query_str}' ===")
        page_token = None
        page_index = 0

        while True:
            page_index += 1
            logging.info(f"[{query_str}] Page #{page_index}, pageToken={page_token!r}")

            # Call search() with retries
            def search_api_call():
                return youtube.search().list(
                    part="snippet",
                    type="video",
                    maxResults=50,
                    order="date",
                    publishedAfter=published_after_str,
                    regionCode="FR",
                    q=query_str,
                    pageToken=page_token
                )

            search_response = youtube_api_call_with_retries(search_api_call, max_retries=3, sleep_seconds=5)
            if not search_response:
                logging.warning(f"[{query_str}] Error calling search().list, skipping the rest.")
                break

            items = search_response.get("items", [])
            logging.info(f"[{query_str}] Page #{page_index} returned {len(items)} videos.")
            if not items:
                logging.info(f"[{query_str}] Empty result -> finishing.")
                break

            total_videos_fetched += len(items)

            # Process videos
            for idx, item in enumerate(items, start=1):
                snippet = item.get("snippet", {})
                video_id = item["id"].get("videoId")
                channel_id = snippet.get("channelId", "")

                logging.info(f"[{query_str} Pg#{page_index} Vid#{idx}] video_id={video_id}, channel_id={channel_id}")

                # 1) Check if we already processed this video
                cur.execute("SELECT 1 FROM processed_videos WHERE video_id=?", (video_id,))
                row = cur.fetchone()
                if row:
                    logging.info(f"    -> Video {video_id} is already in DB, skipping.")
                    continue

                title = snippet.get("title", "")
                description = snippet.get("description", "")

                # Detect language
                text_for_lang = f"{title}\n{description}"
                lang_detected, conf = langid.classify(text_for_lang)
                logging.info(f"    lang={lang_detected}, conf={conf:.4f}")
                if lang_detected != "fr":
                    logging.info("    -> Language != 'fr', skipping.")
                    # Mark video as processed
                    cur.execute("INSERT INTO processed_videos (video_id) VALUES (?)", (video_id,))
                    conn.commit()
                    continue

                # 2) Request channel statistics (channels().list)
                def channels_api_call():
                    return youtube.channels().list(
                        part="statistics",
                        id=channel_id
                    )

                ch_resp = youtube_api_call_with_retries(channels_api_call, max_retries=3, sleep_seconds=5)
                if not ch_resp:
                    logging.warning(f"    -> channels().list returned None, skipping channel {channel_id}.")
                    cur.execute("INSERT INTO processed_videos (video_id) VALUES (?)", (video_id,))
                    conn.commit()
                    continue

                ch_items = ch_resp.get("items", [])
                if not ch_items:
                    logging.info(f"    -> Channel {channel_id} not found in response.")
                    cur.execute("INSERT INTO processed_videos (video_id) VALUES (?)", (video_id,))
                    conn.commit()
                    continue

                stats = ch_items[0].get("statistics", {})
                subs_str = stats.get("subscriberCount", "0")
                try:
                    subs_count = int(subs_str)
                except:
                    subs_count = 0

                logging.info(f"    -> Subscribers: {subs_count}")
                if subs_count >= 50000:
                    logging.info("    -> Too many subscribers (>=50k), skipping.")
                    cur.execute("INSERT INTO processed_videos (video_id) VALUES (?)", (video_id,))
                    conn.commit()
                    continue

                # 3) The channel is suitable, get the handle via Selenium
                logging.info("    -> Calling Selenium for handle...")
                handle = get_handle_from_channel_id_selenium(channel_id, max_retries=3, sleep_seconds=5)
                if not handle:
                    logging.info("    -> Could not get handle, skipping channel.")
                    cur.execute("INSERT INTO processed_videos (video_id) VALUES (?)", (video_id,))
                    conn.commit()
                    continue

                # 4) Check duplicates in df_channels
                if handle in df_channels["channel_handle"].values:
                    logging.info(f"    -> Handle {handle} is already in Excel, skipping.")
                else:
                    # Add a new row
                    logging.info(f"    -> New channel: handle={handle}, subs={subs_count}. Saving to Excel.")
                    df_channels.loc[len(df_channels)] = [handle, subs_count]
                    try:
                        df_channels.to_excel(excel_path, index=False)
                    except PermissionError as pe:
                        logging.error(f"Could not save Excel {excel_path}: {pe}")
                        # If there's a permission issue, data will be lost if the script fails later

                    total_new_channels += 1

                # 5) Mark this video as processed
                cur.execute("INSERT INTO processed_videos (video_id) VALUES (?)", (video_id,))
                conn.commit()

            # Next page
            page_token = search_response.get("nextPageToken")
            if not page_token:
                logging.info(f"[{query_str}] No more pages.")
                break

    # Final
    logging.info("===== RESULT =====")
    logging.info(f"Total videos scanned: {total_videos_fetched} (across all keywords).")
    logging.info(f"New channels added: {total_new_channels}.")

    conn.close()
    logging.info("Script finished.")


if __name__ == "__main__":
    main()

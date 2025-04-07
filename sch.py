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

# Для Selenium
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service as ChromeService
from webdriver_manager.chrome import ChromeDriverManager
from requests.exceptions import ConnectionError as RequestsConnectionError
import urllib3.exceptions

# --------------------------------------------------------------------------
# ФУНКЦИЯ: Повторные попытки для YouTube API (search.list и channels.list)
# --------------------------------------------------------------------------
def youtube_api_call_with_retries(api_func, max_retries=3, sleep_seconds=5):
    """
    Вызывает api_func() (который должен вернуть объект, у которого вызов .execute()),
    делает несколько повторных попыток при HttpError/соединении,
    чтобы обойти временные сбои (ConnectionAbortedError, RemoteDisconnected и т.д.).

    Возвращает response или None, если после max_retries не удалось.
    """
    for attempt in range(1, max_retries + 1):
        try:
            response = api_func().execute()
            return response
        except (HttpError, ConnectionAbortedError, OSError,
                urllib3.exceptions.ProtocolError,
                RequestsConnectionError) as e:
            logging.error(f"[youtube_api_call_with_retries] Попытка {attempt}/{max_retries} -> ошибка: {e}")
            if attempt < max_retries:
                logging.info(f"Подождём {sleep_seconds} сек и повторим...")
                time.sleep(sleep_seconds)
            else:
                logging.error("Лимит повторных попыток исчерпан. Возвращаем None.")
                return None
        except Exception as e2:
            logging.error(f"[youtube_api_call_with_retries] Непредвиденная ошибка: {e2}")
            traceback.print_exc()
            return None
    return None


# --------------------------------------------------------------------------
# ФУНКЦИЯ: Получить handle канала через Selenium (с повторными попытками)
# --------------------------------------------------------------------------
def get_handle_from_channel_id_selenium(channel_id: str,
                                        max_retries=3,
                                        sleep_seconds=5) -> str:
    """
    Открывает https://www.youtube.com/channel/<channel_id> в Selenium,
    ищет селектор:
       div.yt-content-metadata-view-model-wiz__metadata-row
           .yt-content-metadata-view-model-wiz__metadata-row--metadata-row-inline
           span.yt-core-attributed-string--link-inherit-color

    Возвращает что-то вроде '@Evel-901' или None, если не нашлось.
    Пробуем до max_retries раз, если webdriver_manager или сам браузер
    выдают ошибку сети.
    """

    for attempt in range(1, max_retries + 1):
        try:
            return _try_open_channel_and_get_handle(channel_id)
        except (RequestsConnectionError,
                urllib3.exceptions.ProtocolError,
                ConnectionAbortedError,
                Exception) as e:
            # Логируем ошибку, ждём и пробуем ещё
            logging.error(f"[get_handle_from_channel_id_selenium] Попытка {attempt}/{max_retries} -> Ошибка: {e}")
            traceback.print_exc()
            if attempt < max_retries:
                logging.info(f"Подождём {sleep_seconds} сек и повторим открытие Selenium...")
                time.sleep(sleep_seconds)
            else:
                logging.error("Лимит повторных попыток открыть Selenium исчерпан.")
                return None
    return None


def _try_open_channel_and_get_handle(channel_id: str) -> str:
    """
    Реальная логика открытия браузера + поиска handle.
    Вызывается из get_handle_from_channel_id_selenium в цикле повторных попыток.
    """
    logging.info(f"[_try_open_channel_and_get_handle] Старт для channel_id={channel_id}")

    # --- Настройки ChromeOptions ---
    options = webdriver.ChromeOptions()
    # При желании включите безголовый режим:
    # options.add_argument("--headless")

    # Увеличим задержки:
    options.add_argument("--start-maximized")
    options.page_load_strategy = "normal"

    # Если вы уже скачали ChromeDriver вручную, можно указать путь:
    # driver_path = r"C:\path\to\chromedriver.exe"
    # service = ChromeService(executable_path=driver_path)
    # driver = webdriver.Chrome(service=service, options=options)

    # Иначе пытаемся скачать через webdriver_manager:
    service = ChromeService(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=options)

    # Настраиваем большие таймауты
    driver.set_page_load_timeout(120)   # до 120 секунд ждем полной загрузки страницы
    driver.implicitly_wait(30)         # до 30 секунд ждем появления элементов
    logging.info("[Selenium] WebDriver запущен, открываем страницу...")

    try:
        url = f"https://www.youtube.com/channel/{channel_id}"
        driver.get(url)

        # Дополнительная пауза, чтобы точно все подгрузилось (баннер, JS и т.д.)
        time.sleep(15)

        logging.info("[Selenium] Ищем нужный <span>...")

        span_handle = driver.find_element(
            By.CSS_SELECTOR,
            "div.yt-content-metadata-view-model-wiz__metadata-row"
            ".yt-content-metadata-view-model-wiz__metadata-row--metadata-row-inline "
            "span.yt-core-attributed-string--link-inherit-color"
        )

        found_handle = span_handle.text.strip()
        if found_handle:
            logging.info(f"[Selenium] Найден handle: {found_handle}")
            return found_handle
        else:
            logging.warning("[Selenium] Элемент span найден, но текст пуст.")
            return None

    finally:
        logging.info("[Selenium] Закрываем браузер.")
        driver.quit()


# --------------------------------------------------------------------------
# ГЛАВНАЯ ФУНКЦИЯ
# --------------------------------------------------------------------------
def main():
    logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s")

    DEVELOPER_KEY = "AIzaSyDWpLkA5QbQjqz-pfaV03FslYXPGLOn9zg"  # <-- ВСТАВЬТЕ СВОЙ API-KEY!
    youtube = build("youtube", "v3", developerKey=DEVELOPER_KEY)

    # Настройки поиска: за последний год, регион FR
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

    logging.info("=== Настройки ===")
    logging.info(f"Ищем видео, опубликованные после {published_after_str}, regionCode='FR', order='date', maxResults=50.")
    logging.info(f"Французские ключевые слова: {FRENCH_QUERIES}")

    # ----------------------------------------------------------------------
    # Готовим БД (SQLite): "channels_data.db" с таблицей "processed_videos"
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
    # Готовим Excel (channel_info.xlsx)
    # ----------------------------------------------------------------------
    excel_path = "channel_info.xlsx"
    if os.path.exists(excel_path):
        df_channels = pd.read_excel(excel_path, engine="openpyxl")
    else:
        df_channels = pd.DataFrame(columns=["channel_handle", "subscribers"])

    # Проверяем столбцы
    if "channel_handle" not in df_channels.columns:
        df_channels["channel_handle"] = ""
    if "subscribers" not in df_channels.columns:
        df_channels["subscribers"] = 0

    total_videos_fetched = 0
    total_new_channels = 0

    # ----------------------------------------------------------------------
    # Перебираем ключевые слова, листаем выдачу search() по pageToken
    # ----------------------------------------------------------------------
    for query_str in FRENCH_QUERIES:
        logging.info(f"=== Начинаем поиск по запросу '{query_str}' ===")
        page_token = None
        page_index = 0

        while True:
            page_index += 1
            logging.info(f"[{query_str}] Страница #{page_index}, pageToken={page_token!r}")

            # Вызываем search() с ретраями
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
                logging.warning(f"[{query_str}] Ошибка при search().list, пропускаем остаток.")
                break

            items = search_response.get("items", [])
            logging.info(f"[{query_str}] На странице #{page_index} получено {len(items)} видео.")
            if not items:
                logging.info(f"[{query_str}] Пустая выдача -> завершаем.")
                break

            total_videos_fetched += len(items)

            # Обработка видео
            for idx, item in enumerate(items, start=1):
                snippet = item.get("snippet", {})
                video_id = item["id"].get("videoId")
                channel_id = snippet.get("channelId", "")

                logging.info(f"[{query_str} Pg#{page_index} Vid#{idx}] video_id={video_id}, channel_id={channel_id}")

                # 1) Проверяем, не обрабатывали ли уже это видео
                cur.execute("SELECT 1 FROM processed_videos WHERE video_id=?", (video_id,))
                row = cur.fetchone()
                if row:
                    logging.info(f"    -> Видео {video_id} уже в БД, пропускаем.")
                    continue

                title = snippet.get("title", "")
                description = snippet.get("description", "")

                # Определяем язык
                text_for_lang = f"{title}\n{description}"
                lang_detected, conf = langid.classify(text_for_lang)
                logging.info(f"    lang={lang_detected}, conf={conf:.4f}")
                if lang_detected != "fr":
                    logging.info("    -> Язык != 'fr', пропускаем.")
                    # Отмечаем видео как обработанное
                    cur.execute("INSERT INTO processed_videos (video_id) VALUES (?)", (video_id,))
                    conn.commit()
                    continue

                # 2) Запрашиваем статистику канала (channels().list)
                def channels_api_call():
                    return youtube.channels().list(
                        part="statistics",
                        id=channel_id
                    )

                ch_resp = youtube_api_call_with_retries(channels_api_call, max_retries=3, sleep_seconds=5)
                if not ch_resp:
                    logging.warning(f"    -> channels().list вернул None, пропускаем канал {channel_id}.")
                    cur.execute("INSERT INTO processed_videos (video_id) VALUES (?)", (video_id,))
                    conn.commit()
                    continue

                ch_items = ch_resp.get("items", [])
                if not ch_items:
                    logging.info(f"    -> Канал {channel_id} не найден в ответе.")
                    cur.execute("INSERT INTO processed_videos (video_id) VALUES (?)", (video_id,))
                    conn.commit()
                    continue

                stats = ch_items[0].get("statistics", {})
                subs_str = stats.get("subscriberCount", "0")
                try:
                    subs_count = int(subs_str)
                except:
                    subs_count = 0

                logging.info(f"    -> Подписчиков: {subs_count}")
                if subs_count >= 50000:
                    logging.info("    -> Слишком много подписчиков (>=50k), пропускаем.")
                    cur.execute("INSERT INTO processed_videos (video_id) VALUES (?)", (video_id,))
                    conn.commit()
                    continue

                # 3) Канал подходит, получаем handle через Selenium
                logging.info("    -> Вызываем Selenium для handle...")
                handle = get_handle_from_channel_id_selenium(channel_id, max_retries=3, sleep_seconds=5)
                if not handle:
                    logging.info("    -> Не удалось получить handle, пропускаем канал.")
                    cur.execute("INSERT INTO processed_videos (video_id) VALUES (?)", (video_id,))
                    conn.commit()
                    continue

                # 4) Проверяем дубли в df_channels
                if handle in df_channels["channel_handle"].values:
                    logging.info(f"    -> Handle {handle} уже есть в Excel, пропускаем.")
                else:
                    # Добавляем строку
                    logging.info(f"    -> Новый канал: handle={handle}, subs={subs_count}. Записываем в Excel.")
                    df_channels.loc[len(df_channels)] = [handle, subs_count]
                    try:
                        df_channels.to_excel(excel_path, index=False)
                    except PermissionError as pe:
                        logging.error(f"Не удалось сохранить Excel {excel_path}: {pe}")
                        # В этом случае потеряем эти данные, если скрипт упадёт дальше

                    total_new_channels += 1

                # 5) Отмечаем это видео как обработанное
                cur.execute("INSERT INTO processed_videos (video_id) VALUES (?)", (video_id,))
                conn.commit()

            # Следующая страница
            page_token = search_response.get("nextPageToken")
            if not page_token:
                logging.info(f"[{query_str}] Страниц больше нет.")
                break

    # Финал
    logging.info("===== ИТОГ =====")
    logging.info(f"Всего просмотрели {total_videos_fetched} видео (по всем ключевым словам).")
    logging.info(f"Добавлено новых каналов: {total_new_channels}.")

    conn.close()
    logging.info("Скрипт завершён.")


if __name__ == "__main__":
    main()

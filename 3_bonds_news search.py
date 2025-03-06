# 📊 Поиск информации об эмитентах и новостей о компаниях 📊
#
# Этот Python-скрипт автоматически загружает данные об облигациях из Excel-файла,
# получает названия эмитентов через API Московской биржи, затем ищет
# новости по этим компаниям в Google News и сохраняет их в файлы.
#
# Установка зависимостей перед использованием: pip install pandas requests openpyxl feedparser beautifulsoup4 emoji
#
# Автор: Михаил Шардин https://shardin.name/
# Дата создания: 11.02.2025
# Версия: 1.0
#
# Актуальная версия скрипта всегда здесь: https://github.com/empenoso/moex-bond-search-and-analysis
# 

import dataclasses
import os
import time
import re
import requests
import pandas as pd
import feedparser
import urllib.parse
from datetime import datetime
import emoji
import bonds_utils


# 📂 Глобальная переменная для пути к файлу Excel
EXCEL_FILE = "bonds.xlsx"


@dataclasses.dataclass
class NewsItem:
    source: str
    title: str
    date: datetime
    url: str


def load_excel_data(file_path: str = EXCEL_FILE, sheet_name: str = "Исходные данные") -> pd.DataFrame:
    """📂 Загружает данные из Excel и возвращает DataFrame."""
    bonds_utils.log.info("📂 Загружаем данные из Excel...")
    df = pd.read_excel(file_path, sheet_name=sheet_name)
    bonds_utils.log.info(f"✅ Найдено {len(df)} записей")
    return df


def fetch_company_names(df: pd.DataFrame) -> list[str]:
    """🔄 Получает названия компаний по тикерам облигаций."""
    company_names = []
    delay_between_calls = 0.5  # секунды
    for ticker in df.iloc[:, 0]:
        url = f"https://iss.moex.com/iss/securities.json?q={ticker}&iss.meta=off"
        bonds_utils.empty_log.info("")
        bonds_utils.log.info(f"🔍 Обрабатываем тикер: {ticker}")

        try:
            response = requests.get(url)
            response.raise_for_status()
            data = response.json()

            if not data["securities"]["data"]:
                bonds_utils.log.info(f"⚠️ Данные не найдены для {ticker}")
                continue

            emitent_title = data["securities"]["data"][0][8]
            match = re.search(r'"([^"]+)"', emitent_title)
            company_name = match.group(1) if match else emitent_title

            company_names.append(company_name)
            bonds_utils.log.info(f"✅ {emitent_title} → {company_name}")
        
        except (requests.RequestException, IndexError, KeyError) as e:
            bonds_utils.log.info(f"❌ Ошибка при обработке {ticker}: {e}")

        time.sleep(delay_between_calls)

    # 🔄 Удаляем дубликаты, сохраняя порядок
    company_names = list(dict.fromkeys(company_names))
    return company_names


def create_folder() -> str:
    """📂 Создаёт папку для сохранения новостей."""
    current_date = datetime.now().strftime('%Y-%m-%d')
    folder_path = f"news_{current_date}"
    if not os.path.exists(folder_path):
        os.makedirs(folder_path)
    return folder_path


def search_news(company: str) -> list[NewsItem]:
    """🔍 Выполняет поиск новостей по компании."""
    bonds_utils.empty_log.info("")
    bonds_utils.log.info(emoji.emojize(f"🔍 Поиск новостей: {company}"))
    query = urllib.parse.quote(company)
    url = f"https://news.google.com/rss/search?q={query}+when:1y&hl=ru&gl=RU&ceid=RU:ru"
    bonds_utils.log.info(f"📌 Сформирован URL запроса: {url}")
    
    feed: feedparser.FeedParserDict = feedparser.parse(url)
    # TODO: Надо как то типизировать entry
    news_items = [
        NewsItem(
            source=entry.source.title if 'source' in entry else "Google News",
            title=entry.title,
            date=datetime.strptime(entry.published, "%a, %d %b %Y %H:%M:%S %Z"),
            url=entry.link
        )
        for entry in feed.entries
    ]

    bonds_utils.log.info(f"✅ Найдено {len(news_items)} новостей")
    return news_items


def write_news_to_file(folder_path: str, company: str, news: list[NewsItem]) -> None:
    """✍️ Записывает новости в файл."""
    filename = os.path.join(folder_path, f"{company.replace(' ', '_')}.txt")
    with open(filename, 'w', encoding='utf-8') as f:
        f.write(f"📰 Новости для компании {company}\n")
        f.write("=" * 50 + "\n\n")
        
        for item in sorted(news, key=lambda x: x.date, reverse=True):
            f.write(f"📅 Дата: {item.date.strftime('%Y-%m-%d %H:%M')}\n")
            f.write(f"📰 Источник: {item.source}\n")
            f.write(f"📌 Заголовок: {item.title}\n")
            f.write(f"🔗 URL: {item.url}\n")
            f.write("-" * 30 + "\n\n")


def main():
    """🚀 Основная логика выполнения программы."""
    bonds_utils.setup_encoding()
    delay_between_calls = 3  # секунды
    df = load_excel_data()
    company_names = fetch_company_names(df)
    folder_path = create_folder()
    
    for company in company_names:
        news = search_news(company)
        write_news_to_file(folder_path, company, news)
        bonds_utils.log.info(emoji.emojize(f"✍️ Сохранено новостей: {len(news)} для {company}"))
        time.sleep(delay_between_calls)
    
    bonds_utils.empty_log.info("")
    bonds_utils.log.info("🎉 Обработка завершена!")

    # В конце скрипта
    print("\nМихаил Шардин https://shardin.name/\n")
    input("Нажмите любую клавишу для выхода...")

if __name__ == "__main__":
    main()
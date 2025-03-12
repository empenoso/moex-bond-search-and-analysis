from datetime import datetime
import os
import emoji
import urllib.parse

import feedparser
from moex_bond_search_and_analysis.logger import Logger
from moex_bond_search_and_analysis.schemas import NewsItem


def google_search(company: str, log: Logger) -> list[NewsItem]:
    """🔍 Выполняет поиск новостей по компании."""
    log.info(emoji.emojize(f"\n🔍 Поиск новостей: {company}"))
    query = urllib.parse.quote(company)
    url = f"https://news.google.com/rss/search?q={query}+when:1y&hl=ru&gl=RU&ceid=RU:ru"
    log.info(f"📌 Сформирован URL запроса: {url}")
    
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

    log.info(f"✅ Найдено {len(news_items)} новостей")
    return news_items


def write_to_file(folder_path: str, company: str, news: list[NewsItem]) -> None:
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

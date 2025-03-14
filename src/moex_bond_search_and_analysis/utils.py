from datetime import datetime
import os
import sys
import time
from typing import Callable

import humanize

from moex_bond_search_and_analysis.consts import DATETIME_FORMAT


def setup_encoding() -> None:
    # Настройка кодировки для корректного вывода русского текста
    if os.name == "nt":
        sys.stdout.reconfigure(encoding="utf-8")


def measure_method_duration(foo: Callable) -> Callable:
    def wrapper(self, *args, **kwargs):
        start_time = int(time.monotonic())
        self.log.info(
            f"🚀 Функция {foo.__name__} начала работу в {datetime.now().strftime(DATETIME_FORMAT)}."
        )
        result = foo(self, *args, **kwargs)
        duration = humanize.precisedelta(
            int(time.monotonic()) - start_time, minimum_unit="seconds", format="%0.0f"
        )
        self.log.info(
            f"✅ Функция {foo.__name__} закончила работу в {datetime.now().strftime(DATETIME_FORMAT)}."
        )
        self.log.info(f"⏳ Время выполнения {foo.__name__}: {duration}.")
        return result

    return wrapper


def create_news_folder() -> str:
    """📂 Создаёт папку для сохранения новостей."""
    current_date = datetime.now().strftime("%Y-%m-%d")
    folder_path = f"news_{current_date}"
    if not os.path.exists(folder_path):
        os.makedirs(folder_path)
    return folder_path

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

import os
import sys
sys.path.append(f"{os.getcwd()}/src")

from cli import start


if __name__ == "__main__":
    start(3)

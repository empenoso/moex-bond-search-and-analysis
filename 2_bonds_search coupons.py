# 💰 Скачивание и обработка данных о денежном потоке облигаций 💰
#
# Этот Python скрипт автоматически скачивает данные о купонах и выплатах номинала
# через API Московской биржи для списка облигаций из Excel-файла bonds.xlsx и 
# записывает результат обратно в этот же файл.
#
# Установка зависимостей перед использованием: pip install requests openpyxl
#
# Автор: Михаил Шардин https://shardin.name/
# Дата создания: 29.01.2025
# Версия: 1.1
#
# Актуальная версия скрипта всегда здесь: https://github.com/empenoso/moex-bond-search-and-analysis
# 


import os
import sys
sys.path.append(f"{os.getcwd()}/src")

from cli import start


if __name__ == "__main__":
    start(2)

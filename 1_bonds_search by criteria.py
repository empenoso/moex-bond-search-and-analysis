# 🕵️ Поиск ликвидных облигаций Мосбиржи по заданным критериям 🕵️
#
# Этот Python скрипт автоматически выполняет поиск облигаций, соответсвующих заданным
# критериям доходности, цены, дюрации и ликвидности, используя API Московской биржи.
# Результаты поиска, включающие информацию об облигациях и лог действий, 
# записываются в Excel-файл.
#
# Установка зависимостей перед использованием: pip install requests openpyxl
#
# Автор: Михаил Шардин https://shardin.name/
# Дата создания: 14.02.2025
# Версия: 1.3
#
# Актуальная версия скрипта всегда здесь: https://github.com/empenoso/moex-bond-search-and-analysis
#

import os
import sys
sys.path.append(f"{os.getcwd()}/src")

from moex_bond_search_and_analysis.app import App


if __name__ == "__main__":
    App().search_by_criteria()
    print("\nМихаил Шардин https://shardin.name/\n")
    input("Нажмите Enter для выхода...")

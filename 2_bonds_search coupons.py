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

import requests
import openpyxl
from datetime import datetime

# Настройка кодировки для корректного вывода русского текста
import sys
sys.stdout.reconfigure(encoding='utf-8')

# Загружаем Excel-файл
file_path = "bonds.xlsx"
wb = openpyxl.load_workbook(file_path)
sheet_data = wb["Исходные данные"]
sheet_result = wb["Ден.поток"]

# Очищаем лист с результатами
sheet_result.delete_rows(1, sheet_result.max_row)
sheet_result.append(["Название", "Идентификатор", "Дата выплаты", "Денежный поток, ₽ (купон | выплата номинала)"])

# Считываем данные из листа "Исходные данные"
ArraySymbolQuantity = []
for row in sheet_data.iter_rows(min_row=2, max_row=sheet_data.max_row, values_only=True):
    if row[0] and row[1]:  # Проверяем, что данные не пустые
        ArraySymbolQuantity.append(row)

print(f"Считано {len(ArraySymbolQuantity)} облигаций для обработки.")

CashFlow = []

# Обрабатываем каждую облигацию
for ID, number in ArraySymbolQuantity:
    print(f"\nОбрабатываем {ID}, количество: {number} шт.")
    url = f"https://iss.moex.com/iss/statistics/engines/stock/markets/bonds/bondization/{ID}.json?iss.meta=off"
    print(f"Запрос к {url}")
    
    response = requests.get(url)
    json_data = response.json()
    
    # Обработка купонов
    for coupon in json_data.get("coupons", {}).get("data", []):
        name = coupon[1].replace('"', '').replace("'", '').replace("\\", '')
        isin = coupon[0]
        coupon_date = coupon[3]

        # Преобразуем дату в объект datetime
        coupon_datetime = datetime.strptime(coupon_date, "%Y-%m-%d")

        if coupon_datetime > datetime.now():
            value_rub = (coupon[9] or 0) * number
            CashFlow.append([f"{name} (купон 🏷️)", isin, coupon_datetime, value_rub])
            print(f"Добавлен купон: {CashFlow[-1]}")

    # Обработка выплат номинала
    for amort in json_data.get("amortizations", {}).get("data", []):
        name = amort[1].replace('"', '').replace("'", '').replace("\\", '')
        isin = amort[0]
        amort_date = amort[3]

        # Преобразуем дату в объект datetime
        amort_datetime = datetime.strptime(amort_date, "%Y-%m-%d")

        if amort_datetime > datetime.now():
            value_rub = (amort[9] or 0) * number
            CashFlow.append([f"{name} (номинал 💯)", isin, amort_datetime, value_rub])
            print(f"Добавлена выплата номинала: {CashFlow[-1]}")

# Записываем данные в Excel
for row in CashFlow:
    sheet_result.append(row)

# Устанавливаем формат ячеек
for cell in sheet_result["C"][1:]:  # Пропускаем заголовок
    cell.number_format = "DD.MM.YYYY"

for cell in sheet_result["D"][1:]:
    cell.number_format = '# ##0,00 ₽'

# Добавляем запись об обновлении
update_message = f"\nДанные автоматически обновлены {datetime.now().strftime('%d.%m.%Y в %H:%M:%S')}"
sheet_result.append(["", update_message])
print(update_message)

# Сохраняем изменения в файле
wb.save(file_path)
print(f"Файл {file_path} успешно обновлён.")
print("\nМихаил Шардин https://shardin.name/\n")

# В конце скрипта
input("Нажмите любую клавишу для выхода...")
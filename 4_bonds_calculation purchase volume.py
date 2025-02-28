# 💰 Расчет оптимального объема покупки облигаций 💰
#
# Этот Python скрипт автоматически рассчитывает оптимальное количество облигаций для покупки,
# основываясь на доступной сумме денег. Получает актуальные цены и НКД через API Московской биржи
# для списка облигаций из Excel-файла bonds.xlsx и сохраняет результаты расчета
# в новый файл 'bonds_calculation purchase volume.xlsx'.
#
# Функционал:
# - Чтение списка облигаций из исходного Excel-файла
# - Получение актуальных цен и НКД через API Московской биржи
# - Поиск данных за последние 10 дней при отсутствии текущих котировок
# - Равномерное распределение доступной суммы между всеми облигациями
# - Расчет оптимального количества каждой облигации для покупки
# - Сохранение результатов в новый Excel-файл с подробной информацией
#
# Установка зависимостей перед использованием: pip install pandas requests openpyxl
#
# Формат входного файла bonds.xlsx:
# - Лист "Исходные данные"
# - Колонка A: Коды облигаций с Московской биржи
#
# Автор: Михаил Шардин https://shardin.name/
# Дата создания: 16.02.2025
# Версия: 1.0
#
# Актуальная версия скрипта всегда здесь: https://github.com/empenoso/moex-bond-search-and-analysis
#

import pandas as pd
import requests
import json
from datetime import datetime, timedelta

def get_bond_price(security_code):
    """
    # Получение текущей цены облигации и накопленного купонного дохода
    # С попытками получить данные за предыдущие дни при отсутствии текущих данных
    """
    current_date = datetime.now()
    
    for attempt in range(10):
        try_date = current_date - timedelta(days=attempt)
        date_str = try_date.strftime('%Y-%m-%d')
        
        print(f"🔄 Попытка {attempt + 1}: запрос данных за {date_str}")
        
        price_url = f"https://iss.moex.com/iss/history/engines/stock/markets/bonds/boards/TQCB/securities/{security_code}.json?iss.meta=off&iss.json=extended&callback=JSON_CALLBACK&lang=ru&from={date_str}"
        response = requests.get(price_url)
        data = json.loads(response.text.replace('JSON_CALLBACK(', '').rstrip(')'))
        
        if data[1]['history']:
            print(f"✅ Найдены данные за {date_str}")
            close_price = data[1]['history'][0]['CLOSE']
            face_value = data[1]['history'][0]['FACEVALUE']
            current_price = close_price * face_value / 100
            
            nkd_url = f"https://iss.moex.com/iss/engines/stock/markets/bonds/boards/TQCB/securities/{security_code}.json?iss.meta=off&iss.only=securities&lang=ru"
            response = requests.get(nkd_url)
            data = json.loads(response.text)
            accrued_interest = data['securities']['data'][0][7]
            
            return current_price, accrued_interest, date_str
    
    print(f"❌ Не удалось найти данные для {security_code} за последние 10 дней")
    return None, None, None

def calculate_bonds_distribution(available_money):
    """
    # Расчет равномерного распределения средств между облигациями
    """
    print("📊 Чтение списка облигаций из файла Excel...")
    df = pd.read_excel('bonds.xlsx', sheet_name='Исходные данные', usecols='A')
    bonds_list = df.iloc[:, 0].tolist()
    
    # Собираем информацию о всех облигациях
    valid_bonds = []
    for bond in bonds_list:
        print(f"\n🔍 Получение данных для облигации {bond}...")
        price, nkd, date = get_bond_price(bond)
        
        if price is not None:
            valid_bonds.append({
                'bond': bond,
                'price': price,
                'nkd': nkd,
                'total_cost': price + nkd,
                'price_date': date
            })
    
    if not valid_bonds:
        print("❌ Нет доступных облигаций для покупки")
        return []
    
    # Расчет равного распределения денег
    num_bonds = len(valid_bonds)
    money_per_bond = available_money / num_bonds
    print(f"\n💰 Распределение {available_money} руб. между {num_bonds} облигациями")
    print(f"💵 Сумма на каждую облигацию: {money_per_bond:.2f} руб.")
    
    # Расчет количества каждой облигации
    results = []
    for bond_info in valid_bonds:
        num_bonds = int(money_per_bond // bond_info['total_cost'])
        actual_money = num_bonds * bond_info['total_cost']
        
        results.append({
            'bond': bond_info['bond'],
            'quantity': num_bonds,
            'price': bond_info['price'],
            'nkd': bond_info['nkd'],
            'total_cost': bond_info['total_cost'],
            'money_spent': actual_money,
            'price_date': bond_info['price_date']
        })
        
        print(f"\n📈 Облигация {bond_info['bond']}:")
        print(f"   Данные актуальны на: {bond_info['price_date']}")
        print(f"   Цена: {bond_info['price']:.2f} руб.")
        print(f"   НКД: {bond_info['nkd']:.2f} руб.")
        print(f"   Полная стоимость одной облигации: {bond_info['total_cost']:.2f} руб.")
        print(f"   Количество к покупке: {num_bonds} шт.")
        print(f"   Сумма к расходу: {actual_money:.2f} руб.")
    
    # Создание нового DataFrame для результатов
    results_df = pd.DataFrame({
        'Код ценной бумаги': [r['bond'] for r in results],
        'Данные актуальны на': [r['price_date'] for r in results],
        'Цена, руб.': [r['price'] for r in results],
        'НКД, руб.': [r['nkd'] for r in results],
        'Полная стоимость одной облигации, руб.': [r['total_cost'] for r in results],
        'Количество к покупке, шт.': [r['quantity'] for r in results],
        'Сумма к расходу, руб.': [r['money_spent'] for r in results]
    })
    
    # Сохраняем результаты в новый файл
    print("\n📝 Запись результатов в Excel...")
    results_df.to_excel('bonds_calculation purchase volume.xlsx', 
                       sheet_name='Расчет', 
                       index=False)
    print("✅ Результаты сохранены в файл 'bonds_calculation purchase volume.xlsx'")
    
    return results

if __name__ == "__main__":
    available_money = 700000  # Доступная сумма в рублях
    print(f"💵 Доступная сумма: {available_money} руб.")
    results = calculate_bonds_distribution(available_money)
    
    # Вывод итогового распределения средств
    if results:
        total_spent = sum(r['money_spent'] for r in results)
        print(f"\n📊 Итоговое распределение:")
        print(f"Всего потрачено: {total_spent:.2f} руб.")
        print(f"Остаток: {(available_money - total_spent):.2f} руб.")
    
    print("\nМихаил Шардин https://shardin.name/\n")

    # В конце скрипта
    input("Нажмите любую клавишу для выхода...")
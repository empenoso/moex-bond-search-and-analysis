from moex_bond_search_and_analysis.app import App
from moex_bond_search_and_analysis.schemas import SearchByCriteriaConditions
from moex_bond_search_and_analysis.utils import setup_encoding


def start(script_number: None | int = None, search_conditions: SearchByCriteriaConditions | None = None):
    if script_number is None:
        script_number = int(input(
            "1 - Поиск облигаций по критериям\n"
            "2 - Поиск купонов\n"
            "3 - Поиск новостей\n"
            "4 - Подсчет объемов покупки\n"
            "Выберите скрипт: "
        ))

    app = App()
    
    if script_number == 1:
        # Для первого скрипта передаем условия поиска
        app.search_by_criteria(search_conditions=search_conditions)
    elif script_number == 2:
        app.search_coupons()
    elif script_number == 3:
        app.search_news()
    elif script_number == 4:
        app.calc_purchase_volume()
    else:
        print("Выбран неверный номер скрипта.")

    print("\nМихаил Шардин https://shardin.name/\n")
    input("Нажмите Enter для выхода...")


if __name__ == "__main__":
    setup_encoding()
    start()
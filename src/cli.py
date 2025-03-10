from moex_bond_search_and_analysis.app import App
from moex_bond_search_and_analysis.utils import setup_encoding


def start(script_number: None | int = None):
    if script_number is None:
        script_number = int(input(
            "1 - Поиск облигаций по критериям\n"
            "2 - Поиск купонов\n"
            "3 - Поиск новостей\n"
            # "4 - Подсчет объемов покупки\n"
            "Выберите скрипт: "
        ))

    app = App()
    scripts = {
        1: app.search_by_criteria,
        2: app.search_coupons,
        3: app.search_news,
    }
    scripts[script_number]()
    print("\nМихаил Шардин https://shardin.name/\n")
    input("Нажмите Enter для выхода...")


if __name__ == "__main__":
    setup_encoding()
    start()

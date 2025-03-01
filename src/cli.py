from moex_bond_search_and_analysis.app import App


def start():
    script_number = input(
        "1 - Поиск облигаций по критериям\n"
        # "2 - Поиск купонов\n"
        # "3 - Поиск новостей\n"
        # "4 - Подсчет объемов покупки\n"
        "Выберите скрипт: "
    )

    app = App()
    if script_number == "1":
        app.search_by_criteria()


if __name__ == "__main__":
    start()
    print("\nМихаил Шардин https://shardin.name/\n")
    input("Нажмите Enter для выхода...")

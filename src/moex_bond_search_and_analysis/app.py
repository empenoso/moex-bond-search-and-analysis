from datetime import datetime

import emoji
import pandas as pd
import time

from moex_bond_search_and_analysis.moex import MOEX
from moex_bond_search_and_analysis.news import google_search, write_to_file
from moex_bond_search_and_analysis.plugins.excel import ExcelSource
from moex_bond_search_and_analysis.logger import like_print_log
from moex_bond_search_and_analysis.schemas import SearchByCriteriaConditions
from moex_bond_search_and_analysis.utils import create_news_folder, measure_method_duration


class App:
    def __init__(self) -> None:
        self.log = like_print_log
        self.moex = MOEX(log=self.log)
    
    @measure_method_duration
    def search_by_criteria(self):
        search_conditions = SearchByCriteriaConditions()
        moex_search_bonds_result = self.moex.search_bonds(conditions=search_conditions)
        if moex_search_bonds_result:
            output_source = ExcelSource(filename=f"bond_search_{datetime.now().strftime('%Y-%m-%d')}.xlsx")
            output_source.write_search_by_criteria(moex_search_bonds_result, search_conditions, self.moex.log)
            self.log.info(f"\n💾 Результаты записаны в Excel файл: {output_source.filename}")
    
    @measure_method_duration
    def search_coupons(self):
        bounds_source = ExcelSource(filename="bonds.xlsx")
        bond_sheets = bounds_source.load_bonds()
        data_iterator = bond_sheets.data.iter_rows(min_row=2, max_row=bond_sheets.data.max_row, values_only=True)
        bonds = [row for row in data_iterator if row[0] and row[1]]
        self.log.info(f"Считано {len(bonds)} облигаций для обработки.")
        cash_flow = self.moex.process_bonds(bonds=bonds)
        bounds_source.write_bonds(sheets=bond_sheets, cache_flow=cash_flow, log=self.log)

    @measure_method_duration
    def search_news(self):
        delay_between_calls = 3  # секунды
        self.log.info("📂 Загружаем данные из Excel...")
        df = pd.read_excel("bonds.xlsx", sheet_name="Исходные данные")
        self.log.info(f"✅ Найдено {len(df)} записей")
        company_names = self.moex.fetch_company_names(df)
        news_folder_path = create_news_folder()
        for company in company_names:
            news = google_search(company, self.log)
            write_to_file(news_folder_path, company, news)
            self.log.info(emoji.emojize(f"✍️ Сохранено новостей: {len(news)} для {company}"))
            time.sleep(delay_between_calls)

        self.log.info("🎉 Обработка завершена!")

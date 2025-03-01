from datetime import datetime

from moex_bond_search_and_analysis.moex import MOEX
from moex_bond_search_and_analysis.plugins.excel import ExcelSource
from moex_bond_search_and_analysis.logger import like_print_log
from moex_bond_search_and_analysis.schemas import SearchByCriteriaConditions
from moex_bond_search_and_analysis.utils import measure_method_duration


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

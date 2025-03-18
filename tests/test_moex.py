from moex_bond_search_and_analysis.moex import MOEX
from moex_bond_search_and_analysis.logger import like_print_log
from moex_bond_search_and_analysis.schemas import SearchByCriteriaConditions


def test_moex_search_bonds_moex_api_error():
    moex_client = MOEX(log=like_print_log)
    moex_client.API_DELAY = 0
    moex_client.BOARD_GROUPS = [58]
    result = moex_client.search_bonds(conditions=SearchByCriteriaConditions())
    assert result is None
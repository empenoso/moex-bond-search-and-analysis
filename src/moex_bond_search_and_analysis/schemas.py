from datetime import datetime, timedelta
from dataclasses import dataclass, field
from typing import Literal

from openpyxl.workbook import Workbook
from openpyxl.worksheet.worksheet import Worksheet

from moex_bond_search_and_analysis.consts import MONTH_NAMES_RU_SHORT


@dataclass
class Bond:
    name: str = field(metadata={"description": "Полное наименование"})
    secid: str = field(metadata={"description": "Код ценной бумаги"})
    is_qualified_investors: str = field(metadata={"description": "Нужна квалификация?"})
    price: float = field(metadata={"description": "Цена, %"})
    volume: int = field(metadata={"description": "Объем сделок с 15 дней, шт."})
    yield_: float = field(metadata={"description": "Доходность"})
    duration: float = field(metadata={"description": "Дюрация, месяцев"})
    payments_data: dict[str, str] = field(metadata={"description": "Отметки о выплатах"})

    @property
    def as_list(self):
        lst = [self.name, self.secid, self.is_qualified_investors, self.price, self.volume, self.yield_, self.duration]
        lst.extend(
            [self.payments_data.get(month, "") for month in MONTH_NAMES_RU_SHORT]  # Получаем отметки в порядке месяцев
        )
        return lst


@dataclass
class SearchByCriteriaConditions:
    yield_more: int = field(default=15, metadata={"description": "Доходность больше этой цифры"})
    yield_less: int = field(default=40, metadata={"description": "Доходность меньше этой цифры"})
    price_more: int = field(default=70, metadata={"description": "Цена больше этой цифры"})
    price_less: int = field(default=120, metadata={"description": "Цена меньше этой цифры"})
    duration_more: int = field(default=3, metadata={"description": "Дюрация больше этой цифры"})
    duration_less: int = field(default=18, metadata={"description": "Дюрация меньше этой цифры"})
    volume_more: int = field(default=2000, metadata={"description": "Объем сделок в каждый из n дней, шт. больше этой цифры"})
    bond_volume_more: int = field(default=60000, metadata={"description": "Совокупный объем сделок за n дней, шт. больше этой цифры"})
    offer_yes_no: Literal["ДА", "НЕТ"] = field(
        default="ДА",
        metadata={
            "description": (
                "Учитывать, чтобы денежные выплаты были известны до самого погашения?\n"
                "ДА - облигации только с известными цифрами выплаты купонов\n"
                "НЕТ - не важно, пусть в какие-то даты вместо выплаты прочерк"
            )
        }
    )

    @property
    def as_string(self):
        return (
            f"{self.yield_more}% < Доходность < {self.yield_less}%\n"
            f"{self.price_more}% < Цена < {self.price_less}%\n"
            f"{self.duration_more} мес. < Дюрация < {self.duration_less} мес.\n"
            f"Значения всех купонов известны до самого погашения: {self.offer_yes_no}.\n"
            "Объем сделок в каждый из 15 последних дней "
            f"(c {(datetime.now() - timedelta(days=15)).strftime('%d.%m.%Y')}) > {self.volume_more} шт.\n"
            f"Совокупный объем сделок за 15 дней больше {self.bond_volume_more} шт.\n"
            "Поиск в Т0, Т+, Т+ (USD) - Основной режим - безадрес."
        )


@dataclass
class MonthsOfPayments:
    value_rub_null: int
    months_payment_marks: dict[str, str]


@dataclass
class ExcelSheets:
    workbook: Workbook
    data: Worksheet
    result: Worksheet
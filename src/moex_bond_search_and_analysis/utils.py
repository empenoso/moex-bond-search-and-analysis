from datetime import datetime
import os
import sys
import time
from typing import Callable

import humanize

from moex_bond_search_and_analysis.consts import DATETIME_FORMAT
from moex_bond_search_and_analysis.schemas import Bond


def setup_encoding() -> None:
    # Настройка кодировки для корректного вывода русского текста
    if os.name == "nt":
        sys.stdout.reconfigure(encoding="utf-8")


def measure_method_duration(foo: Callable) -> Callable:
    def wrapper(self, *args, **kwargs):
        start_time = int(time.monotonic())
        self.log.info(
            f"🚀 Функция {foo.__name__} начала работу в {datetime.now().strftime(DATETIME_FORMAT)}."
        )
        result = foo(self, *args, **kwargs)
        duration = humanize.precisedelta(
            int(time.monotonic()) - start_time, minimum_unit="seconds", format="%0.0f"
        )
        self.log.info(
            f"✅ Функция {foo.__name__} закончила работу в {datetime.now().strftime(DATETIME_FORMAT)}."
        )
        self.log.info(f"⏳ Время выполнения {foo.__name__}: {duration}.")
        return result

    return wrapper


def create_news_folder() -> str:
    """📂 Создаёт папку для сохранения новостей."""
    current_date = datetime.now().strftime("%Y-%m-%d")
    folder_path = f"news_{current_date}"
    if not os.path.exists(folder_path):
        os.makedirs(folder_path)
    return folder_path

def calculate_XIRR(bond: Bond, cash_flow):
    """
    Calculate the XIRR (Extended Internal Rate of Return) for a series of cash flows.
    
    Args:
        cash_flow (list): List of tuples where each tuple contains (date, cash flow amount).
    
    Returns:
        float: The calculated XIRR value.
    """
    from scipy.optimize import newton

    FEE = 0.3/100  # 0.3% fee T-Bank
    TAX = 0.13  # 13% tax on income

    # buy at current price and spend broker fee and accrued interest
    buy_spend = (bond.price / 100 * bond.lot_value + bond.accrued_interest) * (1 + FEE)
    nice_flow = [[datetime.now().date(), -buy_spend]]

    has_income = bond.lot_value > buy_spend

    for [name, id, date, val, type] in cash_flow:
        if type == "amortization":
            # amortization is taxed if there's income compared to buy_spend
            if has_income:
                # only income is taxed and taxed part is proportional to the amortization value
                val = val - (bond.lot_value - buy_spend) * TAX * (val / bond.lot_value)
    
            nice_flow.append([date.date(), val])
        elif type == "coupon":
            # coupon is taxed
            nice_flow.append([date.date(), val * (1 - TAX)])
        else:
            raise ValueError(f"Unknown cash flow type: {type}")


    print(nice_flow)

    def xirr_function(rate):
        return sum(cf / (1 + rate) ** ((date - nice_flow[0][0]).days / 365.0) for date, cf in nice_flow)

    # Initial guess for the rate
    initial_guess = 0.1
    return newton(xirr_function, initial_guess)

import time

class RateLimiter:
    def __init__(self, min_delay=1.2):
        """
        Initialize the rate limiter.
        
        Args:
            min_delay (float): Minimum delay in seconds between calls (default: 1.2)
        """
        self.min_delay = min_delay
        self.last_call_timestamp = None
    
    def wait_if_needed(self):
        """
        Enforces minimum delay between calls.
        Sleeps if necessary to ensure at least min_delay seconds have passed
        since the last call, then updates the timestamp.
        """

        if self.min_delay == 0: 
            return

        current_time = time.time()
        
        if self.last_call_timestamp is not None:
            if current_time - self.last_call_timestamp < self.min_delay:
                time.sleep(self.min_delay - (current_time - self.last_call_timestamp))
        
        # Update timestamp after potential sleep
        self.last_call_timestamp = time.time()
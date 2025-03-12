import time
import re
from datetime import datetime, timedelta

import pandas as pd
import requests

from moex_bond_search_and_analysis.consts import DATE_FORMAT, MONTH_NAMES_RU_SHORT
from moex_bond_search_and_analysis.logger import Logger
from moex_bond_search_and_analysis.schemas import MonthsOfPayments, SearchByCriteriaConditions, Bond


class MOEX:
    # https://iss.moex.com/iss/engines/stock/markets/bonds/boardgroups/
    BOARD_GROUPS = [58, 193, 105, 77, 207, 167, 245]
    # Переменная для задержки API запросов, лимит в 50 запросов в минуту
    API_DELAY = 1.2

    def __init__(self, log: Logger):
        self.log = log

    def search_bonds(self, conditions: SearchByCriteriaConditions) -> None | list[Bond]:
        """
        Основная функция поиска облигаций по параметрам.
        Выполняет запросы к API Мосбиржи для поиска облигаций, соответствующих заданным критериям.
        Возвращает список найденных облигаций, список сообщений лога и условия поиска. 
        """
        foo_name = "moex_search_bonds"
        bonds = []
        count = 0
        moex_error_counter = 0

        for t in self.BOARD_GROUPS:
            url = (
                f"https://iss.moex.com/iss/engines/stock/markets/bonds/boardgroups/{t}/securities.json"
                "?iss.dp=comma&iss.meta=off&iss.only=securities,marketdata&"
                "securities.columns=SECID,SECNAME,PREVLEGALCLOSEPRICE&marketdata.columns=SECID,YIELD,DURATION"
            )
            self.log.info(f"🔗 {foo_name}. Ссылка поиска всех доступных облигаций группы: {url}.") 

            time.sleep(self.API_DELAY)

            try:
                response = requests.get(url)
                response.raise_for_status()
                json_data = response.json()
            except requests.exceptions.RequestException as e:
                moex_error_counter += 1
                self.log.info(f"⚠️ Ошибка при запросе к API: {e}") 
                continue

            if not json_data or not json_data.get('marketdata') or not json_data['marketdata'].get('data'):
                self.log.info(
                    f'📉 {foo_name}. Нет данных c Московской биржи для группы {t}. Проверьте вручную по ссылке выше.'
                ) 
                continue

            bond_list = json_data['securities']['data']
            count = len(bond_list)
            self.log.info(f'📃 {foo_name}. Всего в списке группы {t}: {count} бумаг.\n') 

            market_data = json_data['marketdata']['data']
            market_data_dict = {item[0]: item for item in market_data if item}  # Создаем словарь для быстрого доступа к данным marketdata по SECID

            for i in range(count):
                # если из-за сетевой ошибки цикл прервался, тогда повтор
                retry_count = 0  # Счётчик попыток
                while retry_count < 5:  # Лимит перезапуска до 5 раз
                    try:
                        bond_name = bond_list[i][1].replace('"', '').replace("'", '')
                        secid = bond_list[i][0]
                        bond_price = bond_list[i][2]

                        bond_market_data = market_data_dict.get(secid)
                        if not bond_market_data:
                            self.log.info(
                                f"❌ {foo_name} в {datetime.now().strftime('%H:%M:%S')}. "
                                f"Строка {i + 1} из {count}: {bond_name} ({secid}): "
                                "Данные о доходности и дюрации отсутствуют."
                            ) 
                            break

                        bond_yield = bond_market_data[1]
                        # кол-во оставшихся месяцев, делим на 30 если есть значение, иначе 0
                        bond_duration = bond_market_data[2] / 30 if bond_market_data[2] else 0
                        bond_duration = round(bond_duration * 100) / 100

                        self.log.info(
                            f"🔎 {foo_name} в {datetime.now().strftime('%H:%M:%S')}. "
                            f"Строка {i + 1} из {count}: {bond_name} ({secid}): "
                            f"цена={bond_price}%, доходность={bond_yield}%, дюрация={bond_duration} мес."
                        ) 

                        condition = (
                            bond_yield is not None and conditions.yield_more <= bond_yield <= conditions.yield_less and
                            bond_price is not None and conditions.price_more <= bond_price <= conditions.price_less and 
                            conditions.duration_more < bond_duration < conditions.duration_less
                        )
                        if condition:
                            self.log.info(
                                f"✅ {foo_name}.   \\-> Условие "
                                f"доходности ({conditions.yield_more} < {bond_yield}% < {conditions.yield_less}), "
                                f"цены ({conditions.price_more} < {bond_price}% < {conditions.price_less}) и "
                                f"дюрации ({conditions.duration_more} < {bond_duration} мес. < {conditions.duration_less}) "
                                f"для {bond_name} прошло."
                            )
                            volume_data = self.search_volume(secid, conditions.volume_more)
                            bond_volume = volume_data['value']
                            self.log.info(
                                f"📊 {foo_name}. \\-> "
                                f"Совокупный объем сделок за n дней: {bond_volume}, а "
                                f"условие {conditions.bond_volume_more} шт."
                            )
                            # lowLiquid: 0 и 1 - переключатели.
                            # ❗ 0 - чтобы оборот был строго больше заданного
                            # ❗ 1 - фильтр оборота не учитывается, в выборку попадают все бумаги, подходящие по остальным параметрам
                            if volume_data['low_liquid'] == 0 and bond_volume > conditions.bond_volume_more:
                                payments_data = self.search_months_of_payments(secid)
                                is_qualified_investors = self.search_is_qualified_investors(secid)
                                bond_instance = Bond(
                                    name=bond_name,
                                    secid=secid,
                                    is_qualified_investors=is_qualified_investors,
                                    price=bond_price,
                                    volume=bond_volume,
                                    yield_=bond_yield,
                                    duration=bond_duration,
                                    payments_data=payments_data.months_payment_marks,  # XXX: похоже тут надо распаковать словарь
                                )
                                if conditions.offer_yes_no == "ДА" and payments_data.value_rub_null == 0:
                                    bonds.append(bond_instance)
                                    self.log.info(f"🗓️ {foo_name}. Для {bond_name} ({secid}) все даты будущих платежей с известным значением выплат.") 
                                    self.log.info(f'⭐ {foo_name}. Результат № {len(bonds)}: {bonds[-1]}.') 
                                elif conditions.offer_yes_no == "НЕТ":
                                    bonds.append(bond_instance)  
                                    self.log.info(f'⭐ {foo_name}. Результат № {len(bonds)}: {bonds[-1]}.\n') 
                                else:
                                    self.log.info(f"🚫 {foo_name}. Облигация {bond_name} ({secid}) в выборку не попадает из-за того, что есть даты когда значения выплат неизвестны.\n") 
                            else:
                                self.log.info(f"💧 {foo_name}. Облигация {bond_name} ({secid}) в выборку не попадает из-за малых оборотов или доступно мало торговых дней.\n") 
                        else:
                            self.log.info(f'⏭️ {foo_name} Пропуск {secid}: не соответствует базовым параметрам.\n') 
                        break

                    except requests.exceptions.RequestException as e:
                        retry_count += 1
                        moex_error_counter += 1
                        self.log.info(f"\n⚠️ Ошибка при обработке строки {i + 1}: {e}.\n🔄 Попытка {retry_count} из 5. Ожидание 60 секунд.\n")
                        time.sleep(60)
                    except Exception as e:
                        retry_count += 1
                        moex_error_counter += 1
                        self.log.info(f"\n🔥 Непредвиденная ошибка при обработке строки {i + 1}: {e}.\n🔄 Попытка {retry_count} из 5. Ожидание 60 секунд.\n")
                        time.sleep(60)

        if not bonds:
            self.log.info(f"📭 {foo_name}. В массиве нет строк.") 
            return None 

        bonds.sort(key=lambda x: x.volume, reverse=True)
        self.log.info(f"📊 {foo_name}. Начало выборки: {bonds[0]}, ...") 
        self.log.info(f"🐞 {foo_name}. Количество ошибок в соединении с Московской биржей: {moex_error_counter}, все данные получены.") 
        return bonds

    def search_volume(self, security_id: str, threshold_value: int) -> dict[str, int]:
        """
        Объем сделок в каждый из n дней больше определенного порога.
        Получает данные об объемах торгов для заданной облигации за последние 15 дней.
        Возвращает словарь с информацией о ликвидности, суммарном объеме и сообщениями лога.
        """
        foo_name = "moex_search_volume"
        now = datetime.now()
        date_request_previous = (now - timedelta(days=15)).strftime(DATE_FORMAT)  # этот день n дней назад
        board_id = self.board_id(security_id)
        if not board_id:
            self.log.info(f"⚠️ Не удалось получить board_id для {security_id}. Поиск объема прерван.") 
            return {'low_liquid': 1, 'value': 0}

        url = (
            f"https://iss.moex.com/iss/history/engines/stock/markets/bonds/boards/{board_id}/securities/{security_id}.json?"
            f"iss.meta=off&iss.only=history&history.columns=SECID,TRADEDATE,VOLUME,NUMTRADES&limit=20&from={date_request_previous}"
        )
        # numtrades - Минимальное количество сделок с бумагой
        # VOLUME - оборот в количестве бумаг (Объем сделок, шт)
        self.log.info(f'🔗 {foo_name}. Ссылка для поиска объёма сделок {security_id}: {url}') 
        try:
            time.sleep(self.API_DELAY)

            response = requests.get(url)
            response.raise_for_status()
            json_data = response.json()
            history_data = json_data['history']['data']

            count = len(history_data)
            volume_sum = 0
            low_liquid = 0
            for i in range(count):
                volume = history_data[i][2]
                volume_sum += volume
                if threshold_value > volume:  # если оборот в конкретный день меньше
                    low_liquid = 1
                    self.log.info(
                        f"📉 {foo_name}. На {i + 1}-й день ({history_data[i][1]}) из {count} "
                        f"оборот по бумаге {security_id} меньше чем {threshold_value}: {volume} шт."
                    ) 
                if count < 6:  # если всего дней в апи на этом периоде очень мало
                    low_liquid = 1
                    self.log.info(
                        f"⚠️ {foo_name}. Всего в АПИ Мосбиржи доступно {count} дней, "
                        f"а надо хотя бы больше 6 торговых дней с {date_request_previous}!"
                    )

            if low_liquid != 1:
                self.log.info(f"📈 {foo_name}. Во всех {count} днях оборот по бумаге {security_id} был больше, чем {threshold_value} шт каждый день.")

            self.log.info(f"📊 {foo_name}. Итоговый оборот в бумагах (объем сделок, шт) за {count} дней: {volume_sum} шт нарастающим итогом.")
            return {'low_liquid': low_liquid, 'value': volume_sum}

        except requests.exceptions.RequestException as e:
            self.log.info(f"⚠️ Ошибка c {security_id} в {foo_name}: {e}")
            return {'low_liquid': 1, 'value': 0}
        except Exception as e:
            self.log.info(f"🔥 Непредвиденная ошибка c {security_id} в {foo_name}: {e}")
            return {'low_liquid': 1, 'value': 0} 

    def board_id(self, security_id: str) -> None | str:
        """
        Узнаем boardid любой бумаги по тикеру.
        Получает board_id для заданной облигации.
        Возвращает board_id или None в случае ошибки.
        """
        foo_name = "moex_board_id"
        url = f"https://iss.moex.com/iss/securities/{security_id}.json?iss.meta=off&iss.only=boards&boards.columns=secid,boardid,is_primary"
        try:
            time.sleep(self.API_DELAY)

            response = requests.get(url)
            response.raise_for_status()
            json_data = response.json()

            board_id_data = json_data['boards']['data']
            primary_board = next((board[1] for board in board_id_data if board[2] == 1), None)  # Находим board_id где is_primary = 1

            if primary_board:
                return primary_board
            else:
                self.log.info(f"⚠️ Не найден primary board_id для {security_id}.") 
                return None

        except requests.exceptions.RequestException as e:
            self.log.info(f"⚠️ Ошибка c {security_id} в {foo_name}: {e}")
            return None
        except Exception as e:
            self.log.info(f"🔥 Непредвиденная ошибка c {security_id} в {foo_name}: {e}")
            return None

    def search_months_of_payments(self, security_id: str) -> MonthsOfPayments:
        """
        Узнаём месяцы, когда происходят выплаты.
        Получает данные о купонных выплатах для заданной облигации.
        Возвращает словарь с информацией о месяцах выплат, наличии неизвестных выплат и months_payment_marks.
        """
        foo_name = "moex_search_months_of_payments"
        url = f"https://iss.moex.com/iss/statistics/engines/stock/markets/bonds/bondization/{security_id}.json?iss.meta=off&iss.only=coupons&start=0&limit=100"
        self.log.info(f'🔗 {foo_name}. Ссылка для поиска месяцев выплат для {security_id}: {url}.') 
        try:
            time.sleep(self.API_DELAY)

            response = requests.get(url)
            response.raise_for_status()
            json_data = response.json()

            coupon_data = json_data['coupons']['data']

            coupon_dates = []
            value_rub_null = 0
            for i in range(len(coupon_data)):
                coupondate = coupon_data[i][3]  # даты купона
                value_rub = coupon_data[i][9]  # сумма выплаты купона
                in_future = datetime.strptime(coupondate, DATE_FORMAT) > datetime.now()
                if in_future:
                    coupon_dates.append(int(coupondate.split("-")[1]))  # Добавляем номер месяца
                    if value_rub is None:
                        value_rub_null += 1

            if value_rub_null > 0:
                self.log.info(f"⚠️ {foo_name}. Для {security_id} есть {value_rub_null} дат(ы) будущих платежей с неизвестным значением выплат.") 

            unique_dates = sorted(list(set(coupon_dates)))  # уникальные значения месяцев и сортировка
            self.log.info(f"🗓️ {foo_name}. Купоны для {security_id} выплачиваются в {unique_dates} месяцы.") 

            months_payment_marks = {}  # Словарь для отметок месяцев
            for month_num in range(1, 13):
                months_payment_marks[MONTH_NAMES_RU_SHORT[month_num-1]] = "✅" if month_num in unique_dates else ""  # Отмечаем месяцы с выплатами

            return MonthsOfPayments(value_rub_null=value_rub_null, months_payment_marks=months_payment_marks)

        except requests.exceptions.RequestException as e:
            self.log.info(f"⚠️ Ошибка c {security_id} в {foo_name}: {e}")
            return  MonthsOfPayments(value_rub_null=0, months_payment_marks={})
        except Exception as e:
            self.log.info(f"🔥 Непредвиденная ошибка c {security_id} в {foo_name}: {e}")
            return  MonthsOfPayments(value_rub_null=0, months_payment_marks={})

    def search_is_qualified_investors(self, security_id: str) -> str:
        """
        Определяем это бумага для квалифицированных инвесторов или нет.
        Получает информацию о необходимости квалификации для покупки облигации.
        Возвращает 'да' или 'нет'.
        """
        foo_name = "moex_search_is_qualified_investors"
        url = f"https://iss.moex.com/iss/securities/{security_id}.json?iss.meta=off&iss.only=description&description.columns=name,title,value"
        self.log.info(f'🔗 {foo_name}. Ссылка для поиска общей информации по {security_id}: {url}') 
        try:

            time.sleep(self.API_DELAY)

            response = requests.get(url)
            response.raise_for_status()
            json_data = response.json()
            description_data = json_data['description']['data']

            is_qualified_investors_data = next((item for item in description_data if item[0] == 'ISQUALIFIEDINVESTORS'), None)
            qual_investor_group_data = next((item for item in description_data if item[0] == 'QUALINVESTORGROUP'), None)

            is_qualified_investors = int(is_qualified_investors_data[2]) if is_qualified_investors_data and is_qualified_investors_data[2] else 0  # По умолчанию 0, если не найдено
            qual_investor_group = qual_investor_group_data[2] if qual_investor_group_data and qual_investor_group_data[2] else "не определена"  # Текст по умолчанию, если не найден

            if is_qualified_investors == 0:
                self.log.info(f"👤 {foo_name}. Для {security_id} квалификация для покупки НЕ нужна.") 
                return 'нет'
            else:
                self.log.info(f"👨‍💼 {foo_name}. {security_id} это бумага для квалифицированных инвесторов категории: \"{qual_investor_group}\"") 
                return 'да'

        except requests.exceptions.RequestException as e:
            self.log.info(f"⚠️ Ошибка c {security_id} в {foo_name}: {e}") 
            return 'ошибка'  # Return some error indicator
        except Exception as e:
            self.log.info(f"🔥 Непредвиденная ошибка c {security_id} в {foo_name}: {e}") 
            return 'ошибка'  # Return some error indicator

    def process_bonds(self, bonds: list[tuple[str | float | datetime | None, ...]]) -> list[list[str]]:
        cash_flow = []
        # Обрабатываем каждую облигацию
        for ID, number in bonds:
            self.log.info("")
            self.log.info(f"Обрабатываем {ID}, количество: {number} шт.")
            url = f"https://iss.moex.com/iss/statistics/engines/stock/markets/bonds/bondization/{ID}.json?iss.meta=off"
            self.log.info(f"Запрос к {url}")
            
            response = requests.get(url)
            json_data = response.json()
            
            assert isinstance(number, (float, int))
            coupons = json_data.get("coupons", {})
            amortizations = json_data.get("amortizations", {})
            cash_flow.extend(self.process_coupons(coupons.get("data", []), coupons.get("columns", []), number))
            cash_flow.extend(self.process_payment(amortizations.get("data", []), amortizations.get("columns", []), number))

        return cash_flow

    def process_coupons(self, coupons: list[tuple[str | int | float, ...]], columns: list[str], number: float | int) -> list[list[str]]:
        # Обработка купонов
        cash_flow = []

        isin_idx = columns.index("isin")
        name_idx = columns.index("name")
        coupondate_idx = columns.index("coupondate")
        value_rub_idx = columns.index("value_rub")

        for coupon in coupons:
            name = str(coupon[name_idx]).replace('"', '').replace("'", '').replace("\\", '')
            isin = coupon[isin_idx]
            coupon_date = coupon[coupondate_idx]

            # Преобразуем дату в объект datetime
            coupon_datetime = datetime.strptime(str(coupon_date), "%Y-%m-%d")

            if coupon_datetime > datetime.now():
                value_rub = float(coupon[value_rub_idx] or 0) * number
                flow = [f"{name} (купон 🏷️)", isin, coupon_datetime, value_rub]
                cash_flow.append(flow)
                self.log.info(f"Добавлен купон: {flow}")

        return cash_flow

    def process_payment(self, amortizations: list[tuple[str | int | float, ...]], columns: list[str], number: float | int) -> list[list[str]]:
        # Обработка выплат номинала
        cash_flow = []

        isin_idx = columns.index("isin")
        name_idx = columns.index("name")
        amortdate_idx = columns.index("amortdate")
        value_rub_idx = columns.index("value_rub")

        for amort in amortizations:
            name = str(amort[name_idx]).replace('"', '').replace("'", '').replace("\\", '')
            isin = amort[isin_idx]
            amort_date = amort[amortdate_idx]

            # Преобразуем дату в объект datetime
            amort_datetime = datetime.strptime(str(amort_date), "%Y-%m-%d")

            if amort_datetime > datetime.now():
                value_rub = float(amort[value_rub_idx] or 0) * number
                flow = [f"{name} (номинал 💯)", isin, amort_datetime, value_rub]
                cash_flow.append(flow)
                self.log.info(f"Добавлена выплата номинала: {flow}")

        return cash_flow

    def fetch_company_names(self, df: pd.DataFrame) -> list[str]:
        """🔄 Получает названия компаний по тикерам облигаций."""
        company_names = []
        delay_between_calls = 0.5  # секунды
        for ticker in df.iloc[:, 0]:
            url = f"https://iss.moex.com/iss/securities.json?q={ticker}&iss.meta=off"
            self.log.info(f"\n🔍 Обрабатываем тикер: {ticker}")

            try:
                response = requests.get(url)
                response.raise_for_status()
                data = response.json()

                if not data["securities"]["data"]:
                    self.log.info(f"⚠️ Данные не найдены для {ticker}")
                    continue

                emitent_title = data["securities"]["data"][0][8]
                match = re.search(r'"([^"]+)"', emitent_title)
                company_name = match.group(1) if match else emitent_title

                company_names.append(company_name)
                self.log.info(f"✅ {emitent_title} → {company_name}")
            
            except (requests.RequestException, IndexError, KeyError) as e:
                self.log.info(f"❌ Ошибка при обработке {ticker}: {e}")

            time.sleep(delay_between_calls)

        # 🔄 Удаляем дубликаты, сохраняя порядок
        company_names = list(dict.fromkeys(company_names))
        return company_names
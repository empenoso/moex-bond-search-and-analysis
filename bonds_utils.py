import logging
import os
import sys


log = logging.getLogger(__name__)
empty_log = logging.getLogger("empty")

log.setLevel(logging.INFO)
empty_log.setLevel(logging.INFO)

handler = logging.StreamHandler(sys.stdout)
empty_handler = logging.StreamHandler(sys.stdout)

formatter = logging.Formatter("%(asctime)s - %(levelname)s - %(message)s")
empty_formater = logging.Formatter("")

handler.setFormatter(formatter)
empty_handler.setFormatter(empty_formater)

log.addHandler(handler)
empty_log.addHandler(empty_handler)


def setup_encoding() -> None:
    # Настройка кодировки для корректного вывода русского текста
    if os.name == "nt":
        sys.stdout.reconfigure(encoding="utf-8")

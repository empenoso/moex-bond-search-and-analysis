import logging
import sys


class Logger:
    def __init__(self, name: str, format: str, store: bool = True):
        self.log = self.__get_logger(name, format)
        self.messages = [] if store else None

    def __get_logger(self, name: str, format: str) -> logging.Logger:
        handler = logging.StreamHandler(sys.stdout)
        formatter = logging.Formatter(format)
        handler.setFormatter(formatter)

        log = logging.getLogger(name)
        log.setLevel(logging.INFO)
        log.addHandler(handler)
        return log

    def info(self, message: str):
        if self.messages is not None:
            if message.startswith("\n"):
                self.messages.append("")
            self.messages.append(message)
            if message.endswith("\n"):
                self.messages.append("")

        self.log.info(message)


# main_log = Logger(name="main", format="%(asctime)s - %(levelname)s - %(message)s", store=True)
like_print_log = Logger(name="main", format="%(message)s", store=True)
# empty_log = Logger(name="empty", format="", store=False)

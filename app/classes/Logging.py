import os
import logging

from packages.config import Config
from packages.dirs import Dirs

class Logging():

    logger = None

    def __init__(self):
        file_name = Config().get('log_file')
        logs_dir = Dirs().get('logs')
        file_path = os.path.join(logs_dir, file_name)

        # создаем регистратор
        self.logger = logging.getLogger('logger')
        self.logger.setLevel(logging.WARNING)

        handler = logging.FileHandler(file_path)
        handler.setLevel(logging.WARNING)

        # строка формата сообщения
        strfmt = '[%(asctime)s] [%(levelname)s] > %(message)s'
        # строка формата времени
        datefmt = '%Y-%m-%d %H:%M:%S'
        # создаем форматтер
        formatter = logging.Formatter(fmt=strfmt, datefmt=datefmt)
        # добавляем форматтер к 'ch'
        handler.setFormatter(formatter)

        self.logger.addHandler(handler)

    def error(self, text):
        self.logger.error(text)

    def warning(self, text):
        self.logger.warning(text)
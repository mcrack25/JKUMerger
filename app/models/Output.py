import os
from packages.dirs import Dirs
from packages.config import Config
from openpyxl import load_workbook, Workbook

class Output():
    file = None

    def __init__(self):
        self.file = Config().get('output_file')
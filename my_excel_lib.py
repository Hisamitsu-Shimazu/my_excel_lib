import openpyxl
import pandas as pd

class MyExcelLib():
    def __init__(self, file_path, file_name):
        self._file_path = file_path
        self._file_name = file_name
        self._book = None

    def create_new_book(self):
        print('create new book.')
        self._book = openpyxl.Workbook()

    def load_book(self):
        print(f'load {self._file_path+self._file_name}')
        self._book = openpyxl.load_workbook(self._file_path + self._file_name)

    def save_book(self):
        print(f'save {self._file_path+self._file_name}')
        self._book.save(self._file_path + self._file_name)
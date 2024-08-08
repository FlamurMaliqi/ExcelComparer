import openpyxl as opx
from ExcelFile import ExcelFile


# Diese Klasse dient dem Einlesen einer Excel-Datei und repräsentiert diese Datei
class ExcelReader:
    def __init__(self, path=str):
        self.path = path
        self.workbook = None
        self.worksheet = None

    # Erstellen einer Excel-Datei, diese kann zur weiteren verarbeitung genutzt werden.
    def init_excel_file(self) -> ExcelFile:
        self._open_excel_file()
        self._get_worksheet()
        file = ExcelFile(self.workbook, self.worksheet, self.path)
        return file

    # Einlesen der Excel-Datei und
    def _open_excel_file(self):
        self.workbook = opx.load_workbook(self.path)

    # Suchen des jeweiligen Worksheets(Arbeitsblatt).
    # Nach Anforderung wird immer nur das erste Arbeitsblatt benötigt, deswegen die hardcode Null.
    def _get_worksheet(self):
        sheets = self.workbook.sheetnames
        self.worksheet = self.workbook[sheets[0]]

    def close_excel_file(self):
        self.workbook.close()

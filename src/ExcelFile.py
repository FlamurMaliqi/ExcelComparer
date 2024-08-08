# Darstellung einer Excel-Datei als einfaches Model.
# Soll als Bindeglied zwischen ExcelReader und ExcelWriter dienen.
class ExcelFile:

    def __init__(self, workbook, worksheet, path):
        self.workbook = workbook
        self.worksheet = worksheet
        self.path = path
        self.max_row = self.worksheet.max_row
        self.max_column = self.worksheet.max_column
        self.heading_line_rows = []
        self.headings = 0
        self.lines_below_headings = []

    def set_workbook(self, workbook):
        self.workbook = workbook

    def get_workbook(self):
        return self.workbook

    def set_worksheet(self, worksheet):
        self.worksheet = worksheet

    def get_worksheet(self):
        return self.worksheet

    # Methode zu testzwecken nicht mehr in Gebrauch.
    def _get_maximum_rows(self):
        rows = 0
        for max_row, row in enumerate(self.worksheet, 1):
            if not all(col.value is None for col in row):
                rows += 1
        return rows + 1
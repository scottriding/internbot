import openpyxl

class Highlighter(object):

    def __init__(self, path_to_xlsx):
        self.workbook = openpyxl.load_workbook(path_to_xlsx)
        self.groups = []
    
    def highlight(self):
        for sheet in self.workbook.worksheets:
            if sheet.title == 'Country':
                row_labels = self.parse_rows(sheet)
                self.parse_columns(sheet)
        
    def parse_rows(self, sheet):
        response_column = sheet['B']
        row_names = []
        for cell in response_column:
            if cell.value is not None and \
               cell.value != 'Sigma' and \
               cell.value not in row_names:
                row_names.append(cell.value)
        return row_names

    def parse_columns(self, sheet):
        group_row = sheet['3']
        for cell in group_row:
            if cell.value is not None and \
               cell.value not in self.groups:
                self.groups.append(cell.value)
import openpyxl
import csv

class RenameTabs(object):

    def __init__(self):
        self.__sheet_names = []

    def rename(self, path_to_xlsx, path_to_tables, path_to_output):
        output_workbook = openpyxl.load_workbook(path_to_xlsx)
        toc_sheet = output_workbook['TOC']
        self.write_table_of_contents(path_to_tables, toc_sheet)
        self.rename_worksheets(output_workbook)
        output_workbook.save(path_to_output + '/Unhighlighted.xlsx') 

    def write_table_of_contents(self, path_to_tables, toc_sheet):
        table_index = 'C'
        question_name = 'D'
        question_prompt = 'E'
        base = 'F'
        iteration = 10
        with open(path_to_tables, 'rb') as csvfile:
            file = csv.DictReader(csvfile, quotechar = '"')
            for table_info in file:
                toc_sheet[table_index + str(iteration)] = table_info['TableIndex']
                sheet_name = table_info['VariableName'].translate(None,"$")
                toc_sheet[question_name + str(iteration)] = sheet_name
                toc_sheet[question_prompt + str(iteration)] = table_info['Title']
                if table_info['Base'] is not '':
                    toc_sheet[base + str(iteration)] = table_info['Base']
                iteration += 1
                self.__sheet_names.append(sheet_name)

    def rename_worksheets(self, workbook):
        index = 0
        for sheet in workbook.worksheets:
            if sheet.title != 'TOC':
                sheet.title = self.__sheet_names[index]
                index += 1
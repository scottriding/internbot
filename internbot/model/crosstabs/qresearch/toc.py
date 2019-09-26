from openpyxl import Workbook
from collections import OrderedDict

class TOC(object):

    def __init__(self):
        self.__workbook = Workbook()

    def build_toc_report(self, survey, path_to_output):
        if self.__workbook.get_sheet_by_name("Sheet") is not None:
            toc_sheet = self.__workbook.get_sheet_by_name("Sheet")
            toc_sheet.title = "TOC"
        else:
            toc_sheet = self.__workbook.create_sheet("TOC")
        self.write_header(toc_sheet)
        self.write_tables(survey.get_questions(), toc_sheet)
        self.__workbook.save(path_to_output)

    def write_header(self, sheet):
        sheet["A1"].value = "Table #"
        sheet["B1"].value = "Question Title"
        sheet["C1"].value = "Base Description"
        sheet["D1"].value = "Base Size"

    def write_tables(self, questions, sheet):
        current_row = 2
        table_iteration = 1
        for question in questions:
            table_no_cell = "A%s" % current_row
            question_title_cell = "B%s" % current_row

            if table_iteration < 10:
                table_name = "Table 0%s" % table_iteration
            else:
                table_name = "Table %s" % table_iteration

            sheet[table_no_cell].value = table_name
            sheet[question_title_cell].value = "%s: %s" % (question.name, question.prompt)
            table_iteration += 1
            current_row += 1



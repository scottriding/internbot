from openpyxl import load_workbook, Workbook
from openpyxl.styles.borders import Border, Side
from openpyxl.styles import PatternFill, Font, Alignment
from collections import OrderedDict
   
class TrendedScoreReport(object):

    def __init__ (self, all_workbook_details, path_to_output, number_of_rounds):
        self.total_count = 0
        rounds = int(number_of_rounds)
        self.configure_cols(rounds)

        # font details
        self.titles_style = Font(name='Arial', size=11, bold=True)
        self.general_style = Font(name='Arial', size=11)
        self.second_row = Font(name='Arial', size=11, bold=True, italic=True)

        # border details
        self.border_all = Border(left=Side(style='thick'), right=Side(style='thick'), top=Side(style='thick'), bottom=Side(style='thick'))
        self.top_border = Border(left=Side(style='thick'), right=Side(style='thick'), top=Side(style='thick'))
        self.middle_border = Border(left=Side(style='thick'), right=Side(style='thick'))
        self.bottom_border = Border(left=Side(style='thick'), right=Side(style='thick'),bottom=Side(style='thick'))

        # fill/highlight details
        self.grey = PatternFill("solid", fgColor="B7B7B7")

        # build reports
        self.build_workbooks(all_workbook_details, path_to_output)

    def configure_cols(self, rounds):
        pass

    def build_workbooks(self, all_workbook_details, path_to_output):
        workbook_names = all_workbook_details.list_workbook_names()
        for workbook_name in workbook_names:
            new_workbook = Workbook()
            self.build_sheets(new_workbook, all_workbook_details.get_workbook(workbook_name))
            save_path = path_to_output + "/" + workbook_name + ".xlsx"
            new_workbook.save(save_path)

    def build_sheets(self, workbook, workbook_details):
        default_sheet = True
        sheet_names = workbook_details.list_sheet_names()
        for sheet_name in sheet_names:
            if default_sheet is True:
                sheet = workbook.get_sheet_by_name("Sheet")
                sheet.title = sheet_name
                default_sheet = False
            else:
                sheet = workbook.create_sheet(sheet_name)
            self.write_sheet(sheet, workbook_details.get_sheet(sheet_name))

    def write_sheet(self, sheet, sheet_details):
        self.write_first_row(sheet)
        current_row_field = 2
        current_row_grouping = 2
        field_names = sheet_details.list_field_names()
        for field_name in field_names:
            location = "A%s" % current_row_field
            sheet[location].value = field_name
            sheet[location].font = self.titles_style
            sheet[location].alignment = Alignment(horizontal="center")
            field_object = sheet_details.get_field(field_name)
            group_count = field_object.grouping_count()
            end = (current_row_field + group_count) - 1
            sheet.merge_cells(start_column=1, end_column=1, start_row=current_row_field, end_row=end)
            current_row_field += group_count
            current_row_grouping = self.write_groupings(sheet, current_row_grouping, field_object)

    def write_first_row(self, current_sheet):
        current_sheet["A1"].value = "Field Name"
        current_sheet["A1"].font = self.titles_style
        current_sheet["A1"].border = self.border_all
        current_sheet["A1"].alignment = Alignment(horizontal="center")

        current_sheet.column_dimensions["A"].width = 35

        current_sheet["B1"].value = "Grouping"
        current_sheet["B1"].font = self.titles_style
        current_sheet["B1"].border = self.border_all
        current_sheet["B1"].alignment = Alignment(horizontal="center")

        current_sheet.column_dimensions["B"].width = 39

        current_sheet["C1"].value = "Count"
        current_sheet["C1"].font = self.titles_style
        current_sheet["C1"].border = self.border_all
        current_sheet["C1"].alignment = Alignment(horizontal="center")

        current_sheet.column_dimensions["C"].width = 11

        current_sheet["D1"].value = "Percent"
        current_sheet["D1"].font = self.titles_style
        current_sheet["D1"].border = self.border_all
        current_sheet["D1"].alignment = Alignment(horizontal="center")

        current_sheet.column_dimensions["D"].width = 11

        current_sheet["E1"].value = "Current Round-Previous Round"
        current_sheet["E1"].font = self.titles_style
        current_sheet["E1"].border = self.border_all
        current_sheet["E1"].alignment = Alignment(horizontal="center", wrap_text = True)

        current_sheet.column_dimensions["E"].width = 15.5

        current_sheet["F1"].value = "Current Round-First Round"
        current_sheet["F1"].font = self.titles_style
        current_sheet["F1"].border = self.border_all
        current_sheet["F1"].alignment = Alignment(horizontal="center", wrap_text = True)

        current_sheet.column_dimensions["F"].width = 14.5

        current_sheet["G1"].value = "Round 1"
        current_sheet["G1"].font = self.titles_style
        current_sheet["G1"].border = self.border_all
        current_sheet["G1"].alignment = Alignment(horizontal="center", wrap_text = True)

        current_sheet.column_dimensions["G"].width = 11

    def write_groupings(self, current_sheet, current_cell, field_object):
        grouping_names = field_object.list_grouping_names()
        middle_border = False
        top_border = True
        if len(grouping_names) > 2:
            middle_border = True
        for grouping_name in grouping_names:
            grouping = field_object.get_grouping(grouping_name)

            location_a = "A%s" % (current_cell)
            location_b = "B%s" % (current_cell)
            location_c = "C%s" % (current_cell)
            location_d = "D%s" % (current_cell)
            location_e = "E%s" % (current_cell)
            location_f = "F%s" % (current_cell)
            location_g = "G%s" % (current_cell)

            if grouping_name in "All Voters":
                self.total_count = grouping.count
                current_sheet[location_a].fill = self.grey
                current_sheet[location_a].font = self.second_row
                current_sheet[location_b].fill = self.grey
                current_sheet[location_b].font = self.second_row
                current_sheet[location_c].fill = self.grey
                current_sheet[location_c].font = self.second_row
                current_sheet[location_d].fill = self.grey
                current_sheet[location_d].font = self.second_row
                current_sheet[location_g].fill = self.grey
                current_sheet[location_g].font = self.second_row

            current_sheet[location_b].value = grouping_name
            current_sheet[location_b].font = self.titles_style
            current_sheet[location_c].value = grouping.count
            current_sheet[location_d].value = float(grouping.count)/float(self.total_count)
            current_sheet[location_e].value = "--%"
            current_sheet[location_f].value = "--%"
            current_sheet[location_g].value = grouping.percent
            current_sheet[location_d].number_format = '0%'
            current_sheet[location_g].number_format = '0%'
            current_sheet[location_c].number_format = '#,##0'

            if top_border is True:
                current_sheet[location_a].border = self.top_border
                current_sheet[location_b].border = self.top_border
                current_sheet[location_c].border = self.top_border
                current_sheet[location_d].border = self.top_border
                current_sheet[location_e].border = self.top_border
                current_sheet[location_f].border = self.top_border
                current_sheet[location_g].border = self.top_border
                top_border = False
            elif middle_border is True:
                current_sheet[location_a].border = self.middle_border
                current_sheet[location_b].border = self.middle_border
                current_sheet[location_c].border = self.middle_border
                current_sheet[location_d].border = self.middle_border
                current_sheet[location_e].border = self.middle_border
                current_sheet[location_f].border = self.middle_border
                current_sheet[location_g].border = self.middle_border

            current_cell += 1

        location_a = "A%s" % (current_cell - 1)
        location_b = "B%s" % (current_cell - 1)
        location_c = "C%s" % (current_cell - 1)
        location_d = "D%s" % (current_cell - 1)
        location_e = "E%s" % (current_cell - 1)
        location_f = "F%s" % (current_cell - 1)
        location_g = "G%s" % (current_cell - 1)
        current_sheet[location_a].border = self.bottom_border
        current_sheet[location_b].border = self.bottom_border
        current_sheet[location_c].border = self.bottom_border
        current_sheet[location_d].border = self.bottom_border
        current_sheet[location_e].border = self.bottom_border
        current_sheet[location_f].border = self.bottom_border
        current_sheet[location_g].border = self.bottom_border

        return current_cell

    def highlight(self, cell):
        darkest_negative = PatternFill("solid", fgColor="B80001")
        medium_dark_negative = PatternFill("solid", fgColor="CD4748")
        medium_negative = PatternFill("solid", fgColor="DF8A8C")
        medium_light_negative = PatternFill("solid", fgColor="EAB9BB")
        lightest_negative = PatternFill("solid", fgColor="F1D2D5")

        darkest_positive = PatternFill("solid", fgColor="02A747")
        medium_dark_positive = PatternFill("solid", fgColor="2AB666")
        medium_positive = PatternFill("solid", fgColor="5AC88B")
        medium_light_positive = PatternFill("solid", fgColor="B6E7CE")
        lightest_positive = PatternFill("solid", fgColor="E6F6F0")

        if cell.value < 0:
            if cell.value >= -0.01:
                cell.fill = lightest_negative
            elif cell.value >= -0.03:
                cell.fill = medium_light_negative
            elif cell.value >= -0.05:
                cell.fill = medium_negative
            elif cell.value >= -0.07:
                cell.fill = medium_dark_negative
            else:
                cell.fill = darkest_negative
        elif cell.value > 0:
            if cell.value <= 0.01:
                cell.fill = lightest_positive
            elif cell.value <= 0.03:
                cell.fill = medium_light_positive
            elif cell.value <= 0.05:
                cell.fill = medium_positive
            elif cell.value <= 0.07:
                cell.fill = medium_dark_positive
            else:
                cell.fill = darkest_positive
        
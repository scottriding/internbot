import csv
import openpyxl

# wb = Workbook()
# del wb["Sheet"]
# for title in ("DropCount_Offenders", "Dropstat_Offenders", "DropCountPerSec_Offenders", "numPktDrops_Offenders"):
#   wb.create_sheet(title)
# 
# 
# for f in filenames:
#     src = DictReader(f)
#     ws = wb[f]
#     ws.append(["Probe_Name", "Recording_Time", "Drop_Count"])
#     for row in src:
#        ws.append(row["Probe_Name"], ["Recording_Time"], ["Drop_Count"])
# 
# wb.save("Drop Offenders.xlsx")

class XLSXMerger(object):

    def merge(self, path_to_folder, path_to_output, path_to_template):
        path_to_csv = path_to_folder + 'Tables to run.csv'
        output_workbook = openpyxl.load_workbook(path_to_template)
        toc_sheet = output_workbook['TOC']
        self.write_table_of_contents(path_to_csv, toc_sheet)
        output_workbook.save(path_to_output + 'test.xlsx')

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
                toc_sheet[question_name + str(iteration)] = table_info['VariableName'].translate(None,"$")
                toc_sheet[question_prompt + str(iteration)] = table_info['Title']
                if table_info['Base'] is not '':
                    toc_sheet[base + str(iteration)] = table_info['Base']
                iteration += 1
        
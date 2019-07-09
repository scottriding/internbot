import os
import fnmatch
import csv
from crosstabs.Format_SPSS_Report.Parse_SPSS_Tables.spss_parser import SPSSParser
from crosstabs.Format_SPSS_Report.Parse_SPSS_Tables.crosstab_report import CrosstabReportWriter

class CrosstabGenerator(object):

    def __init__ (self, path_to_directory):
        parser = SPSSParser()
        table_files = []
        list_of_files = os.listdir(path_to_directory)
        pattern = "*.xlsx"
        for entry in list_of_files:
            if fnmatch.fnmatch(entry, pattern):
                table_split = entry.split(".xlsx")
                to_add = int(table_split[0])
                #to_add = int(entry.translate(None, ".xlsx"))
                table_files.append(to_add)

        table_files.sort()

        for table in table_files:
            filepath = "%s/%s.xlsx" % (path_to_directory, table)
            parser.add_table(filepath)

        self.__crosstab_details = parser.get_tables()

    def write_report(self, path_to_output, resources_filepath):
        report = CrosstabReportWriter(self.__crosstab_details, resources_filepath)
        report.write_report()
        report.save(path_to_output + "/Crosstab Report.xlsx")

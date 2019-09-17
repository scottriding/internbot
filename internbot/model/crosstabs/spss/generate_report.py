from model.crosstabs.format_spss_report.translate_variables import VariableScriptGenerator
from model.crosstabs.format_spss_report.define_tables import TableFileGenerator

from model.crosstabs.format_spss_report.table_files_parser import TablesParser
from model.crosstabs.format_spss_report.table_script_generator import TableScriptGenerator

from model.crosstabs.format_spss_report.spss_parser import SPSSParser
from model.crosstabs.format_spss_report.crosstab_report import CrosstabReportWriter

import os
import fnmatch
import csv


class CrosstabGenerator(object):

    def __init__ (self):
        self.__variables_script = VariableScriptGenerator()
        self.__tables_to_run = TableFileGenerator()

        self.__tables = TablesParser()
        self.__table_script = TableScriptGenerator()

        self.__report_parser = SPSSParser()
        self.__report_writer = CrosstabReportWriter()

    def create_variable_script(self, survey, path_to_output):
        self.__variables_script.define_variables(survey, path_to_output)
        self.__tables_to_run.define_tables(survey, path_to_output)

    def parse_tables(self, path_to_table_file):
        self.__tables.create_tables(path_to_table_file)
        return self.__tables.tables

    def create_table_script(self, tables, banners, embedded_variables, filtering_variable, path_to_output):
        self.__table_script.compile_scripts(tables, banners, embedded_variables, filtering_variable, path_to_output)

    def parse_report_files(self, path_to_directory):
        table_files = []
        list_of_files = os.listdir(path_to_directory)
        pattern = "*.xlsx"
        for entry in list_of_files:
            if fnmatch.fnmatch(entry, pattern):
                table_split = entry.split(".xlsx")
                to_add = int(table_split[0])
                table_files.append(to_add)

        table_files.sort()

        for table in table_files:
            filepath = "%s/%s.xlsx" % (path_to_directory, table)
            self.__report_parser.add_table(filepath)

    def write_report(self, path_to_output, resources_filepath):
        report = CrosstabReportWriter(self.__report_parser.get_tables(), resources_filepath)
        report.write_report()
        report.save(path_to_output)



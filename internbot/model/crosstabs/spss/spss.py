from model.crosstabs.spss import variable_script
from model.crosstabs.spss import tables_to_run
from model.crosstabs.spss import tables
from model.crosstabs.spss import table_script
from model.crosstabs.spss import parser
from model.crosstabs.spss import format_report

import os
import fnmatch
import csv

class SPSS(object):

    def __init__ (self):
        self.__variable_script = variable_script.VariableScript()
        self.__tables_to_run = tables_to_run.TablesToRun()

        self.__tables = tables.Tables()
        self.__table_script = table_script.TableScript()

        self.__parser = parser.Parser()
        self.__formatter = format_report.Formatter()

    def build_variable_script(self, survey, path_to_output):
        self.__variables_script.define_variables(survey, path_to_output)
        self.__tables_to_run.define_tables(survey, path_to_output)

    def build_table_script(self, tables, banners, embedded_variables, filtering_variable, path_to_output):
        self.__table_script.compile_scripts(tables, banners, embedded_variables, filtering_variable, path_to_output)

    def build_spss_model(self, path_to_directory):
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

    def build_spss_report(self, path_to_output, resources_filepath):
        self.__formatter.write_report(self.__parser.get_tables(), resources_filepath)
        self.__formatter.save(path_to_output)



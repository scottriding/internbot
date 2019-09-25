from model.crosstabs.qresearch import qresearch
from model.crosstabs.spss import spss

class Crosstabs(object):

    def __init__(self):
        self.__qresearch = qresearch.QResearch()
        self.__spss = spss.SPSS()

    def build_toc_report(self, survey, path_to_output):
        self.__qresearch.build_toc_report(survey, path_to_output)

    def format_qresearch_report(self, path_to_workbook, resources_filepath, is_qualtrics):
        self.__qresearch.format_qresearch_report(path_to_workbook, resources_filepath, is_qualtrics)

    def save_qresearch_report(self, path_to_output):
        self.__qresearch.save_qresearch_report(path_to_output)

    def build_variable_script(self, survey, path_to_output):
        self.__spss.build_variable_script(survey, path_to_output)

    def build_table_script(self, tables, banners, embedded_variables, filtering_variable, path_to_output):
        self.__spss.build_table_script(tables, banners, embedded_variables, filtering_variable, path_to_output)

    def build_spss_model(self, path_to_directory):
        self.__spss.build_spss_model(path_to_directory)

    def build_spss_report(self, path_to_output, resources_filepath):
        self.__spss.build_spss_report(path_to_output, resources_filepath)

    


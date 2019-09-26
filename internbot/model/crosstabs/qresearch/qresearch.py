from model.crosstabs.qresearch import toc
from model.crosstabs.qresearch import formatter

class QResearch(object):

    def __init__(self):
        self.__toc = toc.TOC()
        self.__formatter = formatter.Formatter()

    def build_toc_report(self, survey, path_to_output):
        self.__toc.build_toc_report(survey, path_to_output)

    def format_qresearch_report(self, path_to_workbook, resources_filepath, is_qualtrics):
        self.__formatter.format_qresearch_report(path_to_workbook, resources_filepath, is_qualtrics)

    def save_qresearch_report(self, path_to_output):
        self.__formatter.save(path_to_output)

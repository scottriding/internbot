from model.crosstabs.format_q_report import compile_toc
from model.crosstabs.format_q_report import format_xlsx

class QReportBuilder(object):

    def __init__(self):
        self.__toc = compile_toc.QTOCCompiler()
        self.__formatter = format_xlsx.QXLSXFormatter()

    def compile_toc(self, survey, path_to_output):
        self.__toc.compile_toc(survey, path_to_output)

    def format_report(self, path_to_workbook, resources_filepath, is_qualtrics):
        self.__formatter.format_report(path_to_workbook, resources_filepath, is_qualtrics)


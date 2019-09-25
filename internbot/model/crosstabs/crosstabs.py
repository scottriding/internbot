from model.crosstabs.qresearch import qresearch

class Crosstabs(object):

    def __init__(self):
        self.__qresearch = qresearch.QResearch()

    def build_toc_report(self, survey, path_to_output):
        self.__qresearch.build_toc_report(survey, path_to_output)

    def format_qresearch_report(self, path_to_workbook, resources_filepath, is_qualtrics):
        self.__qresearch.format_qresearch_report(path_to_workbook, resources_filepath, is_qualtrics)

    def save_qresearch_report(self, path_to_output):
        self.__qresearch.save_qresearch_report(path_to_output)

    


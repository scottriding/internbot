from model.topline.appendix import appendix
from model.topline.document import document
from model.topline.powerpoint import powerpoint

class Topline(object):

    def __init__(self):
        self.__appendix = appendix.Appendix()
        self.__document = document.Document()
        self.__powerpoint = powerpoint.Powerpoint()

    def appendix_document(self, path_to_csv, path_to_output):
        pass

    def appendix_spreadsheet(self, path_to_csv, path_to_output):
        pass

    def basic_qsf_topline_document(self, survey, path_to_freq, path_to_template, path_to_output):
        pass

    def basic_csv_topline_document(self, path_to_freq, path_to_template, path_to_output):
        pass

    def grouped_qsf_topline_document(self, survey, path_to_freq, groups, path_to_template, path_to_output):
        pass

    def grouped_csv_topline_document(self, path_to_freq, groups, path_to_template, path_to_output):
        pass

    def topline_powerpoint(self, path_to_qsf, path_to_freq, path_to_template, path_to_output):
        pass



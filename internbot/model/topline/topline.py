from model.topline.appendix import appendix
from model.topline.document import document
from model.topline.powerpoint import powerpoint

class Topline(object):

    def __init__(self):
        self.__appendix = appendix.Appendix()
        self.__document = document.Document()
        self.__powerpoint = powerpoint.Powerpoint()

    def build_appendix_model(self, path_to_csv):
        self.__appendix.build_appendix_model(path_to_csv)

    def build_appendix_report(self, path_to_output, is_document, image_path, template_path):
        self.__appendix.build_appendix_report(path_to_output, is_document, image_path, template_path)

    def build_document_model(self, path_to_csv, groups, survey):
        self.__document.build_document_model(path_to_csv, groups, survey)

    def build_document_report(self, path_to_template, path_to_output):
        self.__document.build_document_report(path_to_template, path_to_output)

    def build_powerpoint_model(self, path_to_csv, groups, survey):
        self.__powerpoint.build_powerpoint_model(path_to_csv, groups, survey)

    def build_powerpoint_report(self, path_to_template, path_to_output):
        self.__powerpoint.build_powerpoint_report(path_to_template, path_to_output)


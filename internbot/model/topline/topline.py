from model.topline import appendix
from model.topline import document
from model.topline import powerpoint

from model.topline import assigner

class Topline(object):

    def __init__(self):
        self.__appendix = appendix.Appendix()
        self.__document = document.Document()
        self.__powerpoint = powerpoint.Powerpoint()

    def build_appendix_model(self, path_to_csv):
        frequency_assigner = assigner.Assigner(path_to_csv)
        questions = frequency_assigner.question_blocks
        for question in questions:
            print(question)

    def build_appendix_report(self, path_to_output, is_document, image_path, template_path):
        self.__appendix.build_appendix_report(path_to_output, is_document, image_path, template_path)

    def build_document_model(self, path_to_csv, groups, survey):
        frequency_assigner = assigner.Assigner(path_to_csv, groups, survey)
        self.__questions = frequency_assigner.assign()
        for question in self.__questions:
            print(question)

    def build_document_report(self, path_to_template, path_to_output):
        self.__document.build_document_report(path_to_template, path_to_output)

    def build_powerpoint_model(self, path_to_csv, groups, survey):
        frequency_assigner = assigner.Assigner(path_to_csv, groups, survey)
        self.__questions = frequency_assigner.assign()
        for question in self.__questions:
            print(question)

    def build_powerpoint_report(self, path_to_template, path_to_output):
        self.__powerpoint.build_powerpoint_report(path_to_template, path_to_output)
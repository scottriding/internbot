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
        return frequency_assigner.assign()

    def build_appendix_report(self, question_blocks, path_to_output, template_path):
        self.__appendix.build_appendix_report(question_blocks, path_to_output, template_path)

    def build_document_model(self, path_to_csv, groups, survey):
        frequency_assigner = assigner.Assigner(path_to_csv, groups, survey)
        return frequency_assigner.assign()

    def build_document_report(self, question_blocks, path_to_template, path_to_output):
        self.__document.build_document_report(question_blocks, path_to_template, path_to_output)

    def build_powerpoint_model(self, path_to_csv, groups, survey):
        frequency_assigner = assigner.Assigner(path_to_csv, groups, survey)
        return frequency_assigner.assign()

    def pick_template_layout(self, path_to_template):
        return self.__powerpoint.pick_template_layout(path_to_template)

    def build_powerpoint_report(self, question_blocks, layout_index, path_to_output):
        self.__powerpoint.build_powerpoint_report(question_blocks, layout_index, path_to_output)
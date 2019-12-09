from model import model
from view import view

import os
import fnmatch

class Controller(object):

    def __init__(self):
        self.__view = view.View()
        self.__model = model.Model()

        self.__topline_templates = {}
        self.__topline_templates["Y2"] = os.path.join(template_folder, "topline_template.docx")
        self.__topline_templates["QUALTRICS"] = ""
        self.__topline_templates["UT_POLICY"] = os.path.join(template_folder, "utpolicy_top_template.docx")     

        self.__appendix_templates = {}
        self.__appendix_templates["Y2"] = os.path.join(template_folder, "appendix_template.docx")
        self.__appendix_templates["QUALTRICS"] = ""
        self.__appendix_templates["UT_POLICY"] = os.path.join(template_folder, "utpolicy_app_template.docx")

        self.__template_logos = {}
        self.__template_logos["Y2"] = os.path.join(image_folder, "y2_xtabs.png")
        self.__template_logos["QUALTRICS"] = os.path.join(image_folder, "QLogo.png")
        self.__template_logos["UT_POLICY"] = os.path.join(image_folder, "y2_utpol_logo.png")
        self.__template_logos["WHATSAPP"] = os.path.join(image_folder, "whatsapp.png")

    @property
    def view(self):
        return self.__view

    def build_survey(self, path_to_qsf):
        return self.__model.survey(path_to_qsf)

    def rename(self, path_to_xlsx, path_to_toc):
        return self.__model.rename(path_to_xlsx, path_to_toc, image_folder)

    def highlight(self, workbook, path_to_output, is_trended_amazon):
        self.__model.highlight(workbook, path_to_output, is_trended_amazon)

    def build_toc_report(self, survey, path_to_output):
        self.__model.build_toc_report(survey, path_to_output)

    def build_qresearch_report(self, path_to_workbook, template_name):
        self.__model.format_qresearch_report(path_to_workbook, self.__template_logos.get(template_name))

    def save_qresearch_report(self, path_to_output):
        self.__model.save_qresearch_report(path_to_output)

    def build_appendix_model(self, path_to_csv):
        self.__model.build_appendix_model(path_to_csv)

    def build_appendix_report(self, path_to_output, is_document, template_name, other_template=None):
        template_path = self.__appendix_templates.get(template_name)
        image_path = self.__template_logos.get(template_name)
        if other_template is not None:
            template_path = other_template
        self.__model.build_appendix_report(path_to_output, is_document, image_path, template_path)

    def build_document_model(self, path_to_csv, groups, survey):
        self.__model.build_document_model(path_to_csv, groups, survey)

    def build_document_report(self, template_name, path_to_output, template_path=None):
        if template_path is not None:
            path_to_template = template_path
        else:
            path_to_template = self.__topline_templates.get(template_name)
        self.__model.build_document_report(path_to_template, path_to_output)

    def build_powerpoint_model(self, path_to_csv, groups, survey):
        self.__model.build_powerpoint_model(path_to_csv, groups, survey)

    def build_powerpoint_report(self, path_to_template, path_to_output):
        self.__model.build_powerpoint_report(path_to_template, path_to_output)

    def build_scores_model(self, path_to_csv, round, location):
        self.__model.build_scores_model(path_to_csv, round, location)

    def build_scores_report(self, path_to_output):
        self.__model.build_scores_report(path_to_output)

    def build_issues_model(self, path_to_csv, round):
        self.__model.build_issues_model(path_to_csv, round)

    def build_issues_report(self, path_to_output):
        self.__model.build_issues_report(path_to_output)

    def build_trended_model(self, path_to_csv, round):
        self.__model.build_trended_model(path_to_csv, round)

    def build_trended_report(self, path_to_output):
        self.__model.build_trended_report(path_to_output)

if __name__ == '__main__':
    ## directories here only work outside of executable
    template_folder = "resources/templates"
    image_folder = "resources/images"

    controller = Controller()
    controller.view.controller = controller
    controller.view.run()

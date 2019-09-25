from model import model
from view import view

class Controller(object):

    def __init__(self):
        self.__view = view.View()
        self.__model = model.Model()

    @property
    def view(self):
        return self.__view

    def build_survey(self, path_to_qsf):
        return self.__model.survey(path_to_qsf)

    def rename(self, path_to_xlsx, path_to_toc, path_to_output):
        self.__model.rename(path_to_xlsx, path_to_toc, path_to_output)

    def highlight(self, workbook, path_to_output, is_trended_amazon):
        self.__model.highlight(workbook, path_to_output, is_trended_amazon)

    def build_toc_report(self, survey, path_to_output):
        self.__model.build_toc_report(survey, path_to_output)

    def build_qresearch_report(self, path_to_workbook, is_qualtrics):
        self.__model.format_qresearch_report(path_to_workbook, 'resources/images/', is_qualtrics)

    def save_qresearch_report(self, path_to_output):
        self.__model.save_qresearch_report(path_to_output)

    def build_variable_script(self, survey, path_to_output):
        self.__model.build_variable_script(survey, path_to_output)

    def build_table_script(self, tables, banners, embedded_variables, filtering_variable, path_to_output):
        self.__model.build_table_script(tables, banners, embedded_variables, filtering_variable, path_to_output)

    def build_spss_model(self, path_to_directory):
        self.__model.build_spss_model(path_to_directory)

    def build_spss_report(self, path_to_output):
        self.__model.build_spss_report(path_to_output, 'resources/images/')

    def build_appendix_model(self, path_to_csv):
        self.__model.build_appendix_model(path_to_csv)

    def build_appendix_report(self, path_to_output, is_spreadsheet, is_qualtrics):
        self.__model.build_appendix_report(path_to_output, 'resources/images/', 'resources/templates/appendix_template.docx', is_spreadsheet, is_qualtrics)

    def build_document_model(self, path_to_csv, groups, survey):
        self.__model.build_document_model(path_to_csv, groups, survey)

    def build_document_report(self, path_to_output):
        self.__model.build_document_report('resources/templates/topline_template.docx', path_to_output)

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
    controller = Controller()
    controller.view.controller = controller
    controller.view.run()
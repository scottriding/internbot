from model.base import compile
from model.crosstabs import crosstabs
from model.topline import topline
from model.rnc import rnc

class Model(object):

    def __init__(self):
        self.__xtabs = crosstabs.Crosstabs()
        self.__topline = topline.Topline()
        self.__rnc = rnc.RNC()

    def survey(self, path_to_qsf):
        compiler = compile.QSFSurveyCompiler()
        return compiler.compile(path_to_qsf)

    def rename(self, path_to_xlsx, path_to_toc, resources_filepath):
        return self.__xtabs.rename(path_to_xlsx, path_to_toc, resources_filepath)

    def highlight(self, workbook, path_to_output, is_trended_amazon):
        self.__xtabs.highlight(workbook, path_to_output, is_trended_amazon)

    def build_toc_report(self, survey, path_to_output):
        self.__xtabs.build_toc_report(survey, path_to_output)

    def format_qresearch_report(self, path_to_workbook, image_path):
        self.__xtabs.format_qresearch_report(path_to_workbook, image_path)

    def save_qresearch_report(self, path_to_output):
        self.__xtabs.save_qresearch_report(path_to_output)

    def build_variable_script(self, survey, path_to_output):
        self.__xtabs.build_variable_script(survey, path_to_output)

    def build_table_script(self, tables, banners, embedded_variables, filtering_variable, path_to_output):
        self.__xtabs.build_table_script(tables, banners, embedded_variables, filtering_variable, path_to_output)

    def build_spss_model(self, path_to_directory):
        self.__xtabs.build_spss_model(path_to_directory)

    def build_spss_report(self, path_to_output, image_path):
        self.__xtabs.build_spss_report(path_to_output, image_path)

    def build_appendix_model(self, path_to_csv):
        return self.__topline.build_appendix_model(path_to_csv)

    def build_appendix_report(self, question_blocks, path_to_output, is_document, image_path, template_path):
        self.__topline.build_appendix_report(question_blocks, path_to_output, is_document, image_path, template_path)

    def build_document_model(self, path_to_csv, groups, survey):
        return self.__topline.build_document_model(path_to_csv, groups, survey)

    def build_document_report(self, question_blocks, path_to_template, path_to_output):
        self.__topline.build_document_report(question_blocks, path_to_template, path_to_output)

    def build_powerpoint_model(self, path_to_csv, groups, survey):
        return self.__topline.build_powerpoint_model(path_to_csv, groups, survey)

    def pick_template_layout(self, path_to_template):
        return self.__topline.pick_template_layout(path_to_template)

    def build_powerpoint_report(self, question_blocks, layout_index, path_to_output):
        self.__topline.build_powerpoint_report(question_blocks, layout_index, path_to_output)

    def build_scores_model(self, path_to_csv, round, location):
        self.__rnc.build_scores_model(path_to_csv, round, location)

    def build_scores_report(self, path_to_output):
        self.__rnc.build_scores_report(path_to_output)

    def build_issues_model(self, path_to_csv, round):
        self.__rnc.build_issues_model(path_to_csv, round)

    def build_issues_report(self, path_to_output):
        self.__rnc.build_issues_report(path_to_output)

    def build_trended_model(self, path_to_csv, round):
        self.__rnc.build_trended_model(path_to_csv, round)

    def build_trended_report(self, path_to_output):
        self.__rnc.build_trended_report(path_to_output)


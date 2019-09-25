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

    def qresearch_toc(self, path_to_qsf, path_to_output):
        survey = self.survey(path_to_qsf)
        self.__xtabs.qresearch_toc(survey, path_to_output)

    def qresearch_report(self, path_to_xlsx, path_to_output):
        self.__xtabs.qresearch_report(path_to_xlsx, path_to_output)

    def build_appendix_model(self, path_to_csv):
        self.__topline.build_appendix_model(path_to_csv)

    def build_appendix_report(self, path_to_output, path_to_logos, path_to_template, is_spreadsheet, is_qualtrics):
        self.__topline.build_appendix_report(path_to_output, path_to_logos, path_to_template, is_spreadsheet, is_qualtrics)

    def build_document_model(self, path_to_csv, groups, survey):
        self.__topline.build_document_model(path_to_csv, groups, survey)

    def build_document_report(self, path_to_template, path_to_output):
        self.__topline.build_document_report(path_to_template, path_to_output)

    def build_powerpoint_model(self, path_to_csv, groups, survey):
        self.__topline.build_powerpoint_model(path_to_csv, groups, survey)

    def build_powerpoint_report(self, path_to_template, path_to_output):
        self.__topline.build_powerpoint_report(path_to_template, path_to_output)

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


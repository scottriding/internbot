from model.base import compile
from model.crosstabs import crosstabs
from model.topline import topline
from model.rnc import rnc

class Model(object):

    def __init__(self):
        self.__compiler = compile.QSFSurveyCompiler()
        self.__xtabs = crosstabs.Crosstab()
        self.__topline = topline.Topline()
        self.__rnc = rnc.RNC()

    def survey(self, path_to_qsf):
        return self.__compiler.compile(path_to_qsf)

    def qresearch_toc(self, path_to_qsf, path_to_output):
        survey = self.survey(path_to_qsf)
        self.__xtabs.qresearch_toc(survey, path_to_output)

    def qresearch_report(self, path_to_xlsx, path_to_output):
        self.__xtabs.qresearch_report(path_to_xlsx, path_to_output)

    def appendix_document(self, path_to_csv, path_to_output):
        self.__topline.appendix_document(path_to_csv, path_to_output)

    def appendix_spreadsheet(self, path_to_csv, path_to_output):
        self.__topline.appendix_spreadsheet(path_to_csv, path_to_output)

    def basic_qsf_topline_document(self, path_to_qsf, path_to_freq, path_to_template, path_to_output):
        survey = self.survey(path_to_qsf)
        self.__topline.basic_qsf_topline_document(survey, path_to_freq, path_to_template, path_to_output)

    def basic_csv_topline_document(self, path_to_freq, path_to_template, path_to_output):
        self.__topline.basic_csv_topline_document(path_to_freq, path_to_template, path_to_output)

    def grouped_qsf_topline_document(self, path_to_qsf, path_to_freq, groups, path_to_template, path_to_output):
        survey = self.survey(path_to_qsf)
        self.__topline.grouped_qsf_topline_document(survey, path_to_freq, groups, path_to_template, path_to_output)

    def grouped_csv_topline_document(self, path_to_freq, groups, path_to_template, path_to_output):
        self.__topline.grouped_csv_topline_document(path_to_freq, groups, path_to_template, path_to_output)

    def topline_powerpoint(self, path_to_qsf, path_to_freq, path_to_template, path_to_output):
        self.__topline.topline_powerpoint(path_to_qsf, path_to_freq, path_to_template, path_to_output)

    def scores_topline_report(self, path_to_csv, location, round):
        self.__rnc.scores_topline_report(path_to_csv, location, round)

    def issues_trended_report(self, path_to_csv, round):
        self.__rnc.issues_trended_report(path_to_csv, round)

    def trended_scores_reports(self, path_to_csv, round):
        self.__rnc.trended_scores_reports(path_to_csv, round)


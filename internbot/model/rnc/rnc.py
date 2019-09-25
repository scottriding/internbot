from model.rnc.issue_trended import issue_trended
from model.rnc.scores_topline import scores_topline
from model.rnc.trended_score import trended_score

class RNC(object):

    def __init__(self):
        self.__scores_topline = scores_topline.ScoresToplineReportGenerator()
        self.__issues_trended = issue_trended.IssueTrendedReportGenerator()
        self.__trended_scores = trended_score.TrendedScoresReportGenerator()

    def build_scores_model(self, path_to_csv, round, location):
        self.__scores_topline.build_report(path_to_csv, round, location)

    def build_scores_report(self, path_to_output):
        self.__scores_topline.generate_scores_topline(path_to_output)

    def build_issues_model(self, path_to_csv, round):
        self.__issues_trended.build_report(path_to_csv, round)

    def build_issues_report(self, path_to_output):
        self.__issues_trended.generate_issue_trended(path_to_output)

    def build_trended_model(self, path_to_csv, round):
        self.__trended_scores.build_report(path_to_csv, round)

    def build_trended_report(self, path_to_output):
        self.__trended_scores.generate_trended_scores(path_to_output)
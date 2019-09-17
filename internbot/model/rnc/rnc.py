from model.rnc.issue_trended import issue_trended
from model.rnc.scores_topline import scores_topline
from model.rnc.trended_score import trended_score

class RNC(object):

    def __init__(self):
        self.__scores_topline = scores_topline.ScoresToplineReportGenerator()
        self.__issues_trended = issue_trended.IssueTrendedReportGenerator()
        self.__trended_scores = trended_score.TrendedScoresReportGenerator()

    def scores_topline_report(self, path_to_csv, location, round):
        pass

    def issues_trended_report(self, path_to_csv, round):
        pass

    def trended_scores_reports(self, path_to_csv, round):
        pass
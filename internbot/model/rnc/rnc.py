from model.rnc.issue_trended.issue_trended import IssueTrendedReportGenerator
from model.rnc.scores_topline.scores_topline import ScoresToplineReportGenerator
from model.rnc.trended_score.trended_score import TrendedScoresReportGenerator

class RNC(object):

	def build_scores_model(self, path_to_csv, round_number):
		return ScoresToplineReportGenerator().build_models(path_to_csv, round_number)

	def build_scores_report(self, score_details, round_number, location, path_to_output):
		ScoresToplineReportGenerator().build_report(score_details, round_number, location, path_to_output)

	def build_issues_model(self, path_to_csv, round_number):
		return IssueTrendedReportGenerator().build_models(path_to_csv, round_number)

	def build_issues_report(self, issues_details, round_number, path_to_output):
		IssueTrendedReportGenerator().build_report(issues_details, round_number, path_to_output)

	def build_trended_model(self, path_to_csv, round_number):
		return TrendedScoresReportGenerator().build_models(path_to_csv, round_number)

	def build_trended_report(self, trended_details, round_number, path_to_output):
		TrendedScoresReportGenerator().build_report(trended_details, round_number, path_to_output)
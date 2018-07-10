import csv
from issue_trended_score_model import IssueTrendedNetModel
from issue_trended_score_report import IssueTrendedReport
from scores_topline_model import ScoreToplineModels
from scores_topline_report import ScoresToplineReport
from trended_score_model import TrendedModelWorkbooks
from trended_score_report import TrendedScoreReport

class IssueTrendedReportGenerator(object):

    def __init__(self, path_to_csv, round_number):
        self.__models = IssueTrendedNetModel(round_number)
        self.read_csv(path_to_csv)

    def read_csv(self, path_to_csv):
        with open(path_to_csv, 'rb') as csvfile:
            file = csv.DictReader(csvfile, quotechar = '"')
            for model_data in file:
                self.__models.add_model(model_data)

    def generate_issue_trended(self, path_to_output, number_of_rounds):
        report = IssueTrendedReport(self.__models, number_of_rounds)
        report.save(str(path_to_output) + '/trended.xlsx')

class ScoresToplineReportGenerator(object):

    def __init__(self, path_to_csv, round_number):
        self.__models = ScoreToplineModels(round_number)
        self.read_csv(path_to_csv)

    def read_csv(self, path_to_csv):
        with open(path_to_csv, 'rb') as csvfile:
            file = csv.DictReader(csvfile, quotechar = '"')
            for model_data in file:
                self.__models.add_model(model_data)

    def generate_scores_topline(self, path_to_output, report_location, number_of_rounds):
        report = ScoresToplineReport(self.__models, report_location, number_of_rounds)
        report.save(str(path_to_output) + '/scores_topline.xlsx')

class TrendedScoresReportGenerator(object):

    def __init__(self, path_to_csv, round_number):
        self.__workbooks = TrendedModelWorkbooks(round_number)
        self.read_csv(path_to_csv)

    def read_csv(self, path_to_csv):
        with open(path_to_csv, 'rb') as csvfile:
            file = csv.DictReader(csvfile, quotechar = '"')
            for workbook_data in file:
                self.__workbooks.add(workbook_data)

    def generate_trended_scores(self, path_to_output, number_of_rounds):
        report = TrendedScoreReport(self.__workbooks, path_to_output, number_of_rounds)
        
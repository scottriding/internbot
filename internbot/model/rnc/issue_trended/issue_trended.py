from model.rnc.issue_trended import issue_trended_score_model
from model.rnc.issue_trended import issue_trended_score_report

import csv

class IssueTrendedReportGenerator(object):

    def build_report(self, path_to_csv, round_number):
        self.__models = issue_trended_score_model.IssueTrendedNetModel(round_number)
        self.round = round_number
        self.read_csv(path_to_csv)

    def read_csv(self, path_to_csv):
        with open(path_to_csv, 'r') as csvfile:
            file = csv.DictReader(csvfile, quotechar = '"')
            for model_data in file:
                self.__models.add_model(model_data)

    def generate_issue_trended(self, path_to_output):
        report = issue_trended_score_report.IssueTrendedReport(self.__models, self.round)
        report.save(str(path_to_output))
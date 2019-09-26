from model.rnc.trended_score import trended_score_model
from model.rnc.trended_score import trended_score_report

import csv

class TrendedScoresReportGenerator(object):

    def build_report(self, path_to_csv, round_number):
        self.__workbooks = trended_score_model.TrendedModelWorkbooks(round_number)
        self.round = round_number
        self.read_csv(path_to_csv)

    def read_csv(self, path_to_csv):
        with open(path_to_csv, 'r') as csvfile:
            file = csv.DictReader(csvfile, quotechar = '"')
            for workbook_data in file:
                self.__workbooks.add(workbook_data)

    def generate_trended_scores(self, path_to_output):
        report = trended_score_report.TrendedScoreReport(self.__workbooks, path_to_output, self.round)


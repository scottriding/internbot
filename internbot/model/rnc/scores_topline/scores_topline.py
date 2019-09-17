from model.rnc.scores_topline import scores_topline_model
from model.rnc.scores_topline import scores_topline_report

import csv

class ScoresToplineReportGenerator(object):

    def build_report(self, path_to_csv, round_number):
        self.__models = scores_topline_model.ScoreToplineModels(round_number)
        self.round = round_number
        self.read_csv(path_to_csv)

    def read_csv(self, path_to_csv):
        with open(path_to_csv, 'r') as csvfile:
            file = csv.DictReader(csvfile, quotechar = '"')
            for model_data in file:
                self.__models.add_model(model_data)

    def generate_scores_topline(self, path_to_output, report_location):
        report = scores_topline_report.ScoresToplineReport(self.__models, report_location, self.round)
        report.save(str(path_to_output))

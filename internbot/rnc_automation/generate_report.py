import csv
from model import Models
from rnc_report import RNCReport

class ScoresToplineGenerator(object):
    def __init__(self, path_to_csv):
        self.__models = Models()
        self.read_csv(path_to_csv)

    def read_csv(self, path_to_csv):
        with open(path_to_csv, 'rb') as csvfile:
            file = csv.DictReader(csvfile, quotechar = '"')
            for model_data in file:
                self.__models.add(model_data)

    def generate_rnc_topline(self, path_to_output):
        report = RNCReport(self.__models)
        report.save(str(path_to_output) + '/scores_topline.xlsx')
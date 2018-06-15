import csv
from topline_model import ToplineModels
from trended_model import TrendedModels
from rnc_topline import RNCToplineReport
from rnc_trended import RNCTrendedReport

class ScoresToplineGenerator(object):
    def __init__(self, path_to_csv):
        self.__models = ToplineModels()
        self.read_csv(path_to_csv)

    def read_csv(self, path_to_csv):
        with open(path_to_csv, 'rb') as csvfile:
            file = csv.DictReader(csvfile, quotechar = '"')
            for model_data in file:
                self.__models.add(model_data)

    def generate_rnc_topline(self, statename, path_to_output):
        report = RNCToplineReport(self.__models, statename)
        report.save(str(path_to_output) + '/scores_topline.xlsx')

class ScoresTrendedGenerator(object):
    def __init__(self, path_to_csv):
        self.__models = TrendedModels()
        self.read_csv(path_to_csv)

    def read_csv(self, path_to_csv):
        with open(path_to_csv, 'rb') as csvfile:
            file = csv.DictReader(csvfile, quotechar = '"')
            for model_data in file:
                self.__models.add(model_data)

    def generate_rnc_trended(self, path_to_output):
        report = RNCTrendedReport(self.__models)
        report.save(str(path_to_output) + '/trended.xlsx')
        
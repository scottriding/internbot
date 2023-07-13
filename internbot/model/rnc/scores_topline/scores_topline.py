from model.rnc.scores_topline import scores_topline_model
from model.rnc.scores_topline import scores_topline_report

import csv

class ScoresToplineReportGenerator(object):

	def build_models(self, path_to_csv, round_number):
		self.check_input(path_to_csv, round_number)
		models = self.read_csv(path_to_csv, round_number)
		return models
	
	def check_input(self, csv_path, round_number):
		self.check_input_names(round_number, open(csv_path, encoding="utf8"))

	def check_input_names(self,round_number, utf8_data, **kwargs):
		csv_reader = csv.DictReader(utf8_data, **kwargs)
		for row in csv_reader:
			if row.get("Model") is None:
				raise ValueError(f'Missing column: Model')
			if row.get("Survey question reference") is None:
				raise ValueError(f'Missing column: Survey question reference')
			if row.get("Variable") is None:
				raise ValueError(f'Missing column: Field Name')

			for current in range(1,round_number+1):
					freq_col = "Round %s Frequency" % current
					freq_tow_col = "Round %s TOW Frequency" % current
					date_col = "Round %s Date" % current
					
					if row.get(freq_col) is None:
						raise ValueError(f'Missing column: {freq_col}')
					
					if row.get(freq_tow_col) is None:
						raise ValueError(f'Missing column: {freq_tow_col}')
					
					if row.get(date_col) is None:
						raise ValueError(f'Missing column: {date_col}')
			break

	def read_csv(self, path_to_csv, round_number):
		models = scores_topline_model.ScoreToplineModels(round_number)
		with open(path_to_csv, 'r') as csvfile:
			file = csv.DictReader(csvfile, quotechar = '"')
			for model_data in file:
				models.add_model(model_data)
		return models

	def build_report(self, models, round_number, location, path_to_output):
		report = scores_topline_report.ScoresToplineReport(models, location, round_number)
		report.save(str(path_to_output))

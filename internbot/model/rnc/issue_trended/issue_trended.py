from model.rnc.issue_trended import issue_trended_score_model
from model.rnc.issue_trended import issue_trended_score_report

import csv

class IssueTrendedReportGenerator(object):

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
			if row.get("Field Name") is None:
				raise ValueError(f'Missing column: Field Name')
			if row.get("Grouping") is None:
				raise ValueError(f'Missing column: Grouping')
			if row.get("Count") is None:
				raise ValueError(f'Missing column: Count')

			for current in range(1,round_number+1):
					freq_col = "Round %s Frequency" % current
					date_col = "Round %s Date" % current
					
					if row.get(freq_col) is None:
						raise ValueError(f'Missing column: {freq_col}')
					
					if row.get(date_col) is None:
						raise ValueError(f'Missing column: {date_col}')
			break

	def read_csv(self, path_to_csv, round_number):
		models = issue_trended_score_model.IssueTrendedModels(round_number)
		with open(path_to_csv, 'r') as csvfile:
			file = csv.DictReader(csvfile, quotechar = '"')
			for model_data in file:
				models.add_model(model_data)
		return models

	def build_report(self, models, round_number, path_to_output):
		report = issue_trended_score_report.IssueTrendedReport(models, round_number)
		report.save(str(path_to_output))
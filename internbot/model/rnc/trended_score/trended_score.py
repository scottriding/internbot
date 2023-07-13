from model.rnc.trended_score import trended_score_model
from model.rnc.trended_score import trended_score_report

import csv

class TrendedScoresReportGenerator(object):

	def build_models(self, path_to_csv, round_number):
		self.check_input(path_to_csv, round_number)
		return self.read_csv(path_to_csv)

	def check_input(self, csv_path, round_number):
		self.check_input_names(round_number, open(csv_path, encoding="utf8"))

	def check_input_names(self,round_number, utf8_data, **kwargs):
		csv_reader = csv.DictReader(utf8_data, **kwargs)
		for row in csv_reader:
			if row.get("Workbook") is None:
				raise ValueError(f'Missing column: Workbook')
			if row.get("Sheet") is None:
				raise ValueError(f'Missing column: Sheet')
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
		workbooks = trended_score_model.TrendedModelWorkbooks(round_number)
		with open(path_to_csv, 'r') as csvfile:
			file = csv.DictReader(csvfile, quotechar = '"')
			for workbook_data in file:
				workbooks.add(workbook_data)
		return workbooks

	def build_report(trended_details, round_number, path_to_output):
		report = trended_score_report.TrendedScoreReport(trended_details, path_to_output, round_number)


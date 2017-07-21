from compile import QSFSurveyCompiler
from topline_report import ToplineReport
import csv

class ReportGenerator(object):
	def __init__(self, path_to_qsf):
		self.__questions = []
		self.qsf = QSFSurveyCompiler()
		self.survey = self.qsf.compile(path_to_qsf)
	
	def generate_report(self, path_to_csv, path_to_template, path_to_output):
		with open(path_to_csv, 'rb') as csvfile:
			file = csv.DicReader(csvfile, quotechar = '"')
			for question_data in file:
				matching_question = self.find_question(question_data['name'], self.survey)
				matching_response = self.find_response(question_data['response'], matching_question)
				self.add_frequency(matching_response, question_data['frequency'])
		report = ToplineReport(self.__questions, path_to_template)
		report.save(path_to_output)

	def find_question(self, question_to_find, survey):
        matching_question = survey.blocks.find_question_by_name(question_to_find)
    	if matching_question not in self.__questions:
        	self.__questions.append(matching_question)
    	return matching_question	
    	
   def find_response(self, response_to_find, question):
    	responses = question.responses
    	matching_response = next((response for response in responses if response.response == response_to_find), None)
    	return matching_response 	

	def add_frequency(self, response, frequency):
    	response.frequency = frequency
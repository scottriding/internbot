from topline_report import ToplineReport
from topline_ppt import ToplinePPT
import csv
import re

class ReportGenerator(object):
    def __init__(self, survey):
        self.__questions = []
        self.survey = survey

    def generate_docx(self, path_to_csv, path_to_template, path_to_output):
        self.generate_report(path_to_csv)
        report = ToplineReport(self.survey.get_questions(), path_to_template)
        report.save(path_to_output)        

    def generate_ppt(self, path_to_template, path_to_output):
        report = ToplinePPT(self.survey.get_questions(), path_to_template)
        report.save(path_to_output)

    def generate_report(self, path_to_csv):
        with open(path_to_csv, 'rb') as csvfile:
            file = csv.DictReader(csvfile, quotechar = '"')
            for question_data in file:
                matching_question = self.find_question(question_data['name'], self.survey)
                matching_response = self.find_response(question_data['response'], matching_question)
                if matching_response is not None:
                    self.add_frequency(matching_response, question_data['frequency'])
            
    def find_question(self, question_to_find, survey):
        matching_question = survey.blocks.find_question_by_name(question_to_find)
        if matching_question.type == 'Composite':
            matching_question = self.find_sub_question(matching_question, question_to_find)
        if matching_question not in self.__questions:
            self.__questions.append(matching_question)
        return matching_question

    def find_sub_question(self, composite_question, question_to_find):
        for sub_question in composite_question.questions:
            if re.match(sub_question.name, question_to_find):
                absolute_match = sub_question
        return absolute_match

    def find_response(self, response_to_find, question):
        responses = question.responses
        matching_response = matching_response = next((response for response in responses if response.code == response_to_find), None)
        return matching_response
        
    def add_frequency(self, response, frequency):
        response.frequency = frequency
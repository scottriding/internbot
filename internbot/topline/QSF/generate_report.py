from topline_report import ToplineReport
from topline_ppt import ToplinePPT
from topline_appendix import ToplineAppendix
import csv
import re

class ReportGenerator(object):
    def __init__(self, survey):
        self.__questions = survey.get_questions()
        self.survey = survey

    def generate_basic_topline(self, path_to_csv, path_to_template, path_to_output):
        self.assign_frequencies(path_to_csv)
        report = ToplineReport(self.__questions, path_to_template)
        report.save(str(path_to_output) + '/basic_topline.docx')

    def generate_full_topline(self, path_to_csv, path_to_template, path_to_output, path_to_appendix):
        self.assign_text_responses(path_to_appendix)
        self.generate_basic_topline(path_to_csv, path_to_template, path_to_output)
        self.assign_frequencies(path_to_csv)
        open_ended_questions = [question for question in self.__questions \
                                if question.text_entry == True]
        report = ToplineAppendix()
        report.write_with_topline(open_ended_questions, str(path_to_output) + '/basic_topline.docx')
        report.save(str(path_to_output) + '/full_topline.docx')

    def generate_appendix(self, path_to_template, path_to_appendix, path_to_output):
        self.assign_text_responses(path_to_appendix)
        report = ToplineAppendix()
        open_ended_questions = [question for question in self.__questions \
                                if question.text_entry == True]
        report.write_independent(open_ended_questions, path_to_template)
        report.save(str(path_to_output) + '/appendix.docx')

    def generate_ppt(self, path_to_template, path_to_output):
        report = ToplinePPT(self.__questions, path_to_template)
        report.save(str(path_to_output) + '/topline.pptx')

    def assign_text_responses(self, path_to_appendix):
        with open(path_to_appendix, 'rb') as appendix_file:
            file = csv.DictReader(appendix_file, quotechar = '"')
            for response in file:
                matching_question = self.find_question(response['variable'], self.survey)
                if matching_question is not None:
                    matching_question.add_text_response(response['label'])
                    matching_question.text_entry = True

    def assign_frequencies(self, path_to_csv):
        with open(path_to_csv, 'rb') as csvfile:
            file = csv.DictReader(csvfile, quotechar = '"')
            for question_data in file:
                matching_question = self.find_question(question_data['variable'], self.survey)
                if matching_question is not None:
                    matching_response = self.find_response(question_data['label'], matching_question)
                    if matching_response is not None:
                        self.add_frequency(matching_response, question_data['percent'])
                        self.add_n(matching_question, question_data['n'])
                
    def find_question(self, question_to_find, survey):
        matching_question = survey.blocks.find_question_by_name(question_to_find)
        if matching_question is None:
            return None
        if matching_question.parent == 'CompositeQuestion':
            matching_question = self.find_sub_question(matching_question, question_to_find)
        return matching_question

    def find_sub_question(self, composite_question, question_to_find):
        for sub_question in composite_question.questions:
            if re.match(sub_question.name, question_to_find):
                absolute_match = sub_question
        return absolute_match

    def find_response(self, response_to_find, question):
        if response_to_find == 'NA':
            question.add_NA()
            return question.get_NA()
        responses = question.responses
        matching_response = next((response for response in responses if response.response == response_to_find), None)
        if response_to_find == 'On':
            matching_response = next((response for response in responses if response.response == '1'), None)
        return matching_response

    def add_frequency(self, response, frequency):
        response.frequency = frequency

    def add_n(self, question, n):
        current_n = question.n
        question.n = current_n + n

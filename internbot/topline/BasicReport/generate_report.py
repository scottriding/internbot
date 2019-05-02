from csv_question import CSVQuestions
from csv_topline_report import CSVToplineReport
from qsf_topline_report import QSFToplineReport
import csv
import re
from collections import OrderedDict

class ReportGenerator(object):

    def __init__(self, path_to_freqs, round_no = 1, survey = None):
        question_data = self.unicode_dict_reader(open(path_to_freqs))
        self.__frequencies = []
        if survey is not None:
            self.__survey = survey
            self.__questions = survey.get_questions()
            self.assign_frequencies(question_data, round_no)
        else:
            self.__questions = CSVQuestions()
            self.create_questions(question_data, round_no)
        
    def unicode_dict_reader(self, utf8_data, **kwargs):
        csv_reader = csv.DictReader(utf8_data, **kwargs)
        for row in csv_reader:
            if row['variable'] != "":
                yield {unicode(key, 'iso-8859-1'):unicode(value, 'iso-8859-1') for key, value in row.iteritems()}

    def create_questions(self, question_data, round_no):
        for question in question_data:
            self.__questions.add(question, round_no)

    def assign_frequencies(self, question_data, round_no):
        for row in question_data:
            question_name = row["variable"]
            response_label = row["label"]
            question_display = row["display"]
            matching_question = self.find_question(question_name)
            if matching_question is not None:
                matching_response = self.find_response(row["label"], matching_question)
                if matching_response is not None:
                    self.add_frequency(matching_response, row)
                    self.add_n(matching_question, row)
                    if question_display != "":
                        self.add_display_logic(matching_question, question_display)

    def generate_topline(self, path_to_template, path_to_output, years):
        if self.__survey is not None:
            report = QSFToplineReport(self.__survey.get_questions(), path_to_template, years)
        else:
            report = CSVToplineReport(self.__questions, path_to_template, years)
        report.save(str(path_to_output) + '/topline_report.docx')

    def find_question(self, question_to_find):
        matching_question = self.__survey.blocks.find_question_by_name(question_to_find)
        if matching_question is None:
            return None
        elif matching_question.parent == "CompositeQuestion":
            matching_question = self.find_sub_question(matching_question, question_to_find)
        return matching_question

    def find_sub_question(self, composite_question, question_to_find):
        for sub_question in composite_question.questions:
            if re.match(sub_question.name, question_to_find):
                absolute_match = sub_question
        return absolute_match

    def find_response(self, response_to_find, matching_question):
        if response_to_find == 'NA':
            matching_question.add_NA()
            return matching_question.get_NA()
        responses = matching_question.responses
        matching_response = next((response for response in responses if response.response == response_to_find), None)
        if response_to_find == 'On':
            matching_response = next((response for response in responses if response.response == '1'), None)
        return matching_response

    def add_frequency(self, matching_response, frequency_data):
        round_col = "percent"
        self.__frequencies = []
        if frequency_data[round_col] is not None:
            self.__frequencies.append(frequency_data[round_col])
        else:
        	iteration = 1
        	round_int = int(round_no)
        	while iteration <= round_int:
        		round_col = "percent %s" % iteration
        		if frequency_data[round_col] != '':
        			self.__frequencies.append(frequency_data[round_col])
        		iteration += 1

        matching_response.frequencies = self.__frequencies

    def add_n(self, matching_question, question_data):
        current_n = matching_question.n
        matching_question.n = current_n + int(question_data["n"])

    def add_display_logic(self, matching_question, question_display):
        matching_question.display_logic = question_display
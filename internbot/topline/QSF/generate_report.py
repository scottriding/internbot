from topline_report import ToplineReport
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

    def unicode_dict_reader(self, utf8_data, **kwargs):
        csv_reader = csv.DictReader(utf8_data, **kwargs)
        for row in csv_reader:
            if row['variable'] != "":
                yield {unicode(key, 'iso-8859-1'):unicode(value, 'iso-8859-1') for key, value in row.iteritems()}

    def assign_frequencies(self, path_to_csv):
        question_data = self.unicode_dict_reader(open(path_to_csv))
        for row in question_data:
            matching_question = self.find_question(row["variable"], self.survey)
            if matching_question is not None:
                matching_response = self.find_response(row["label"], matching_question)
                if matching_response is not None:
                    self.add_frequency(matching_response, row["percent"])
                    self.add_n(matching_question, row["n"])
                    if row["display"] != "":
                        self.add_display_logic(matching_question, row["display"])
                
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
        question.n = current_n + int(n)

    def add_display_logic(self, question, display):
        question.display_logic = display

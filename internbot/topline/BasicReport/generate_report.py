from topline.BasicReport.csv_question import CSVQuestions
from topline.BasicReport.csv_topline_report import CSVToplineReport
from topline.BasicReport.qsf_topline_report import QSFToplineReport
import csv
import re
from collections import OrderedDict

class ReportGenerator(object):

    def __init__(self, path_to_freqs, years=[], survey = None):
        question_data = self.unicode_dict_reader(open(path_to_freqs))
        self.__frequencies = []
        self.headers = []
        self.__survey = survey
        if survey is not None:
            self.__survey = survey
            self.__questions = survey.get_questions()
            for question in self.__questions:
                print(question.name)
            self.assign_frequencies(question_data, years)
        else:
            self.__questions = CSVQuestions()
            self.create_questions(question_data, years)
        
    def unicode_dict_reader(self, utf8_data, **kwargs):
        csv_reader = csv.DictReader(utf8_data, **kwargs)
        self.headers = csv_reader.fieldnames
        for row in csv_reader:
            if row['variable'] != "":
                yield {key: value for key, value in row.items()}

    def create_questions(self, question_data, years):
        for question in question_data:
            self.__questions.add(question, self.headers, years)

    def assign_frequencies(self, question_data, years):
        for row in question_data:
            question_name = row["variable"]
            response_label = row["label"]
            question_stat = row["stat"]
            if 'display logic' in self.headers:
                question_display = row["display logic"]
            else:
                question_display = ""

            matching_question = self.find_question(question_name)
            if matching_question is not None:
                # print("MATCHING QUESTION: "+matching_question.name+" ROW: "+question_name)
                matching_response = self.find_response(row["value"], matching_question)
                if matching_response is not None:
                    self.add_frequency(matching_response, row, years)
                    self.add_n(matching_question, row)
                    self.add_display_logic(matching_question, question_display)
                    self.add_stat(matching_question, question_stat)

    def generate_topline(self, path_to_template, path_to_output, years):
        if self.__survey is not None:
            report = QSFToplineReport(self.__survey.get_questions(), path_to_template, years)
        else:
            report = CSVToplineReport(self.__questions, path_to_template, years)
        report.save(str(path_to_output))

    def find_question(self, question_to_find):
        matching_question = self.__survey.blocks.find_question_by_name(question_to_find)
        if matching_question is None:
            print("\nCould not match question "+question_to_find+
                  " from CSV to a question in the QSF.\n        "
                  "     *This data will need to be input manually.*\n")
            return None
        elif matching_question.parent == "CompositeQuestion":
            matching_question = self.find_sub_question(matching_question, question_to_find)
        return matching_question

    def find_sub_question(self, composite_question, question_to_find):
        absolute_match = None
        
        for sub_question in composite_question.questions:
            if re.match(sub_question.name, question_to_find):
                absolute_match = sub_question
        return absolute_match

    def find_response(self, response_to_find, matching_question):
        if response_to_find == 'NA':
            matching_question.add_NA()
            return matching_question.get_NA()
        responses = matching_question.responses

        matching_response = next((response for response in responses if response.code == response_to_find), None)
        if response_to_find == 'On':
            matching_response = next((response for response in responses if response.response == '1'), None)

        if matching_response is None:
            print("\nCould not match response " +response_to_find+ " from " +
                  matching_question.name + " from CSV to a question in the QSF.\n"
                                           "             *This data will need to be input manually.*\n")
        return matching_response

    def add_frequency(self, matching_response, frequency_data, years):
        self.__frequencies = OrderedDict()
        if len(years) > 0:
            for year in years:
                round_col = "result %s" % year
                if frequency_data[round_col] != "":
                    self.__frequencies[year] = frequency_data[round_col]
        else:
            round_col = "result"
            self.__frequencies[0] = frequency_data[round_col]
        matching_response.frequencies = self.__frequencies

    def add_n(self, matching_question, question_data):
        current_n = matching_question.n
        matching_question.n = current_n + int(question_data["n"])

    def add_display_logic(self, matching_question, question_display):
        matching_question.display_logic = question_display

    def add_stat(self, matching_question, stat):
        matching_question.stat = stat

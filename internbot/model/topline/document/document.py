from model.base import survey
from model.base import block
from model.base import question
from model.topline.document import topline_report

import csv
import re
from collections import OrderedDict

class Document(object):

    def build_document_model(self, path_to_freqs, groups=[], survey=None):
        assigner = FrequencyAssigner(path_to_freqs, groups, survey)
        self.__questions = assigner.assign()
        for question in self.__questions:
            print(question)

    def build_document_report(self, path_to_template, path_to_output):
        pass

class FrequencyAssigner(object):

    def __init__(self, path_to_freqs, groups, inputted_survey):
        self.__groups = groups
        self.__frequency_data = self.unicode_dict_reader(open(path_to_freqs))

        if inputted_survey is None:
            inputted_survey = survey.Survey("CSV Survey")
            default_block = block.Block("Default")
            questions = self.create_questions(self.unicode_dict_reader(open(path_to_freqs)))
            default_block.questions = questions
            inputted_survey.add_block(default_block)
        self.__question_blocks = inputted_survey.blocks

    def unicode_dict_reader(self, utf8_data, **kwargs):
        csv_reader = csv.DictReader(utf8_data, **kwargs)
        for row in csv_reader:
            if row['variable'] != "":
                yield {key: value for key, value in row.items()}

    def create_questions(self, question_data):
        questions = question.Questions()
        for response_row in question_data:
            question_name = response_row['variable']
            question_prompt = response_row['prompt']
            matching_question = questions.find_by_name(question_name)
            if matching_question:
                matching_question.add_response(response_row['label'], response_row['value'])
            else:
                new_question = question.Question()
                new_question.name = question_name
                new_question.prompt = question_prompt
                new_question.add_response(response_row['label'], response_row['value'])
                questions.add(new_question)
        return questions

    def assign(self):
        for response_row in self.__frequency_data:
            question_name = response_row['variable']
            matching_question = self.find_question(question_name)
            if matching_question:
                response_value = response_row['value']
                matching_response = self.find_response(response_label, response_value, matching_question)
                if matching_response:
                    self.add_frequency(matching_response, response_row)
        return self.__question_blocks

    def find_question(self, question_to_find):
        matching_question = self.__question_blocks.find_question_by_name(question_to_find)
        if matching_question.parent == 'CompositeQuestion':
            matching_question = self.find_sub_question(matching_question, question_to_find)
        return matching_question

    def find_sub_question(self, composite_question, question_to_find):
        absolute_match = None
        for sub_question in composite_question.questions:
            if re.match(sub_question.name, question_to_find):
                absolute_match = sub_question
        return absolute_match

    def find_response(self, response_label, response_value, matching_question):
        responses = matching_question.responses

        matching_response = next((response for response in responses if response.value == response_value), None)
        if matching_response is None:
             matching_response = next((response for response in responses if response.label == response_label), None)

        ## hotspot edge case
        if response_label == 'On':
            matching_response = next((response for response in responses if response.label == '1'), None)

        return matching_response

    def add_frequency(self, matching_response, response_row):
        stat = response_row['stat']

        if len(self.__groups) > 0:
            for group in self.__groups:
                result_col = "result %s" % group
                n_col = "n %s" % group

                matching_response.add_frequency(response_row[result_col], response_row[n_col], stat, group)
        else:
            result_col = "result"
            n_col = "n"
            matching_response.add_frequency(response_row[result_col], response_row[n_col], stat)
        

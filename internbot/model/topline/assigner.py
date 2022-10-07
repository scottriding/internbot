from model.base import survey
from model.base import block
from model.base import question

import csv
import re
from collections import OrderedDict

class Assigner(object):

    def __init__(self, path_to_freqs, groups=[], inputted_survey=None):
        self.__groups = groups
        self.__survey = inputted_survey

        ### we're going to create a survey object from the frequency file ###
        self.__match_survey = survey.Survey("Frequencies to match")
        match_block = block.Block("Default")
        match_block.questions = self.create_questions(self.unicode_dict_reader(open(path_to_freqs)))
        self.__match_survey.add_block(match_block)

    def unicode_dict_reader(self, utf8_data, **kwargs):
        csv_reader = csv.DictReader(utf8_data, **kwargs)
        for row in csv_reader:
            if row['variable'] != "":
                yield {key: value for key, value in row.items()}

    def create_questions(self, question_data):
        questions = question.Questions()
        for response_row in question_data:
            matching_question = questions.find_by_name(response_row["variable"])
            if not matching_question:
                matching_question = question.Question()
                matching_question.name = response_row["variable"]
                matching_question.prompt = response_row["prompt"]
                matching_question.type = "MC"

                questions.add(matching_question)

            if response_row.get("value"):
                response = matching_question.add_response(response_row["label"], response_row["value"])
                self.add_frequency(response, response_row)
            else:
                response = matching_question.add_response(response_row["label"])
        return questions

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

    def assign(self):
        if not self.__survey:
            return self.__match_survey.blocks

        ### here's where we match the qsf to the frequency file ###
        for qsf_block in self.__survey.blocks:
            for qsf_question in qsf_block.questions:
                if qsf_question.parent is "CompositeQuestion":
                    for subquestion in qsf_question.questions:
                        name_match = self.__match_survey.blocks.find_question_by_name(subquestion.name)
                        self.compare_responses(subquestion, name_match)   
                else:
                    name_match = self.__match_survey.blocks.find_question_by_name(qsf_question.name)
                    self.compare_responses(qsf_question, name_match)

        return self.__survey.blocks

    def compare_responses(self, qsf_question, name_match):
        if name_match:
            qsf_question.assigned = name_match.name
            for qsf_response in qsf_question.responses:
                matching_response = next((response for response in name_match.responses if response.label == qsf_response.label), None)
                if not matching_response:
                    try:
                        matching_response = matching_response = next((response for response in name_match.responses if int(response.value) == int(qsf_response.value)), None)
                    except:
                        pass

                if matching_response:
                    qsf_response.assigned = matching_response
                    qsf_response.frequencies = matching_response.frequencies
                else:
                    qsf_response.assigned = "None"
        else:
            qsf_question.assigned = "None"


        

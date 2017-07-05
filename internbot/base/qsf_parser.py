import json
import re
from survey import Survey
from block import Blocks, Block
from question import Questions, Question

class QSFParser(object):

    def __init__(self):
        self.survey_parser = QSFSurveyParser()
        self.blocks_parser = QSFBlocksParser()
        self.questions_parser = QSFQuestionsParser()

    def parse(self, path_to_qsf):
        qsf_json = self.parse_json(path_to_qsf)
        survey = self.survey_parser.parse(qsf_json['SurveyEntry'])
        blocks = self.blocks_parser.parse(self.find_element('BL', qsf_json))
        for block in blocks:
            survey.add_block(block)
        questions = self.questions_parser.parse(self.find_elements('SQ', qsf_json))
        survey.add_questions(questions)
        return survey

    def parse_json(self, path_to_qsf):
        with open(path_to_qsf) as file:
            qsf_file = json.load(file)
        return qsf_file

    def find_element(self, element_name, qsf_json):
        return next(iter(self.find_elements(element_name, qsf_json)), None)

    def find_elements(self, element_name, qsf_json):
        elements = qsf_json['SurveyElements']
        return [element for element in elements if element['Element'] == element_name]

class QSFSurveyParser(object):

    def parse(self, survey_element):
        return Survey(survey_element['SurveyName'])

class QSFBlocksParser(object):

    def parse(self, blocks_element):
        blocks = Blocks()
        for block_element in blocks_element['Payload']:
            if block_element['Description'] != 'Trash / Unused Questions':
                block = Block(block_element['Description'])
                for question_id in block_element['BlockElements']:
                    if question_id['Type'] == 'Question':
                        block.assign_id(question_id['QuestionID'])
                blocks.add(block)
        return blocks

class QSFQuestionsParser(object):

    def __init__(self):
        self.matrix_parser = QSFQuestionsMatrixParser()

    def parse(self, question_elements):
        questions = []
        for question_element in question_elements:
            question_payload = question_element['Payload']
            question = Question()
            question.id = question_payload['QuestionID']
            question.name = question_payload['DataExportTag']
            question.prompt = question_payload['QuestionText']
            question.type = question_payload['QuestionType']

            if question_payload.get('DynamicChoices'):
                question.has_carry_forward_responses = True
                carry_forward_locator = question_payload['DynamicChoices']['Locator']
                carry_forward_match = re.match('q://(QID\d+).+', carry_forward_locator)
                question.carry_forward_question_id = carry_forward_match.group(1)

            if question.type == 'Matrix':
                matrix_questions = self.matrix_parser.parse(question_payload)
                for question in matrix_questions:
                    questions.append(question)
            else:
                question.subtype = question_payload['Selector']
                if question_payload.get('Choices') and len(question_payload['Choices']) > 0:
                    question.response_order = question_payload['ChoiceOrder']
                    for code, response in question_payload['Choices'].iteritems():
                        question.add_response(response['Display'], code)
                questions.append(question)

        questions = self.carry_forward_responses(questions)
        return questions

    def carry_forward_responses(self, questions):
        dynamic_questions = [question for question in questions if question.has_carry_forward_responses == True]
        for dynamic_question in dynamic_questions:
            matching_question = next((question for question in questions if question.id == dynamic_question.carry_forward_question_id))
            dynamic_question.response_order = matching_question.response_order
            for response in matching_question.responses:
                dynamic_question.add_response(response.response, response.code)
        return questions


class QSFQuestionsMatrixParser(object):

    def parse(self, question_payload):
        matrix_questions = []
        prompts = question_payload['Choices']
        responses = question_payload['Answers']
        for code, prompt in prompts.iteritems():
            question = Question()
            question.id = '%s_%s' % (str(question_payload['QuestionID']), code)
            question.subtype = question_payload['SubSelector']
            question.name = '%s_%s' % (str(question_payload['DataExportTag']), code)
            question.prompt = prompt['Display']
            question.response_order = question_payload['AnswerOrder']
            for code, response in responses.iteritems():
                question.add_response(response['Display'], code)
            matrix_questions.append(question)
        return matrix_questions


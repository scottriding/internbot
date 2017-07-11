import json
import re
from survey import Survey
from block import Blocks, Block
from question import Questions, Question, CompositeQuestion

class QSFParser(object):

    def __init__(self):
        self.survey_parser = QSFSurveyParser()
        self.blockflow_parser = QSFBlockFlowParser()
        self.blocks_parser = QSFBlocksParser()
        self.questions_parser = QSFQuestionsParser()

    def parse(self, path_to_qsf):
        qsf_json = self.parse_json(path_to_qsf)
        survey = self.survey_parser.parse(qsf_json['SurveyEntry'])
        block_ids = self.blockflow_parser.parse(self.find_element('FL', qsf_json))
        blocks = self.blocks_parser.parse(self.find_element('BL', qsf_json))
        blocks.sort(block_ids)
        questions = self.questions_parser.parse(self.find_elements('SQ', qsf_json))
        survey = self.compile_survey(survey, blocks, questions)
        return survey
        
    def parse_json(self, path_to_qsf):
        with open(path_to_qsf) as file:
            qsf_file = json.load(file)
        return qsf_file
        
    def compile_survey(self, survey, blocks, questions):
        for block in blocks:
            survey.add_block(block)
        survey.add_questions(questions)
        return survey

    def find_element(self, element_name, qsf_json):
        return next(iter(self.find_elements(element_name, qsf_json)), None)

    def find_elements(self, element_name, qsf_json):
        elements = qsf_json['SurveyElements']
        return [element for element in elements if element['Element'] == element_name]

class QSFSurveyParser(object):

    def parse(self, survey_element):
        return Survey(survey_element['SurveyName'])
        
class QSFBlockFlowParser(object):
    
    def parse(self, flow_element):
        block_ids = []
        flow_payload = flow_element['Payload']['Flow']
        for block in flow_payload:
            block_ids.append(block['ID'])
        return block_ids    
            

class QSFBlocksParser(object):

    def parse(self, blocks_element):
        blocks = Blocks()
        for block_element in blocks_element['Payload']:
            if block_element['Type'] != 'Trash':
                block = Block(block_element['Description'])
                block.blockid = block_element['ID']
                for question_id in block_element['BlockElements']:
                    if question_id['Type'] == 'Question':
                        block.assign_id(question_id['QuestionID'])
                blocks.add(block)
        return blocks

class QSFQuestionsParser(object):

    def __init__(self):
        self.matrix_parser = QSFQuestionsMatrixParser()
        self.response_parser = QSFResponsesParser()

    def parse(self, question_elements):
        questions = []
        for question_element in question_elements:
            question_payload = question_element['Payload']
            question = self.parse_question_basics(question_payload)
            
            if question_payload.get('DynamicChoices') and question.type != 'Matrix':
                self.assign_carry_forward(question, question_payload)

            if question.type == 'Matrix':
                matrix_question = self.matrix_parser.parse(question_payload)
                questions.append(matrix_question)
            else:
                self.response_parser.parse(question, question_payload, question_element)
                questions.append(question)

        questions = self.response_parser.carry_forward_responses(questions)
        questions = self.carry_forward_prompts(questions)
        return questions
        
    def parse_question_basics(self, question_payload):
        question = Question()
        question.id = question_payload['QuestionID']
        question.name = question_payload['DataExportTag']
        question.prompt = question_payload['QuestionText']
        question.type = question_payload['QuestionType']
        return question
        
    def assign_carry_forward(self, question, question_payload):
        question.has_carry_forward_responses = True
        carry_forward_locator = question_payload['DynamicChoices']['Locator']
        carry_forward_match = re.match('q://(QID\d+).+', carry_forward_locator)
        question.carry_forward_question_id = carry_forward_match.group(1)
            
    def carry_forward_prompts(self, questions):
        dynamic_questions = [question for question in questions if question.has_carry_forward_prompts == True]
        carried_forward_questions = []
        for dynamic_question in dynamic_questions:
            matching_questions = [] # Find all the questions that have the same prefix as my question_id carry forward
            for matching_question in matching_questions:
                question = Question()
                question.prompt = matching_question.prompt
                question.type = dynamic_question.type
                question.subtype = dynamic_question.subtype
                question.has_carry_forward_prompts = dynamic_question.has_carry_forward_prompts
                question.carry_forward_question_id = dynamic_question.carry_forward_question_id
                question.id = '%s_%s' % (dynamic_question.id, matching_question)
                question.name = '%s_%s' % (dynamic_question.id, matching_question)
                carried_forward_questions.append(question)
        return questions

class QSFQuestionsMatrixParser(object):

    def parse(self, question_payload):
        matrix_question = CompositeQuestion()
        prompts = question_payload['Choices']
        responses = question_payload['Answers']
        if question_payload.get('DynamicChoices') is None:
            for code, prompt in prompts.iteritems():
                question = Question()
                question.id = '%s_%s' % (str(question_payload['QuestionID']), code)
                question.code = code
                question.type = question_payload['QuestionType']
                question.subtype = question_payload['SubSelector']
                if question_payload['ChoiceDataExportTags']:
                    question.name = question_payload['ChoiceDataExportTags'][code]    
                else:
                    question.name = '%s_%s' % (str(question_payload['DataExportTag']), code)
                question.prompt = prompt['Display']
                question.response_order = question_payload['AnswerOrder']
                for code, response in responses.iteritems():
                    question.add_response(response['Display'], code)
                matrix_question.add_question(question)
                matrix_question.id = question_payload['QuestionID']
                matrix_question.question_order = question_payload['ChoiceOrder']
        else:
            question = CompositeQuestion()
            question.id = question_payload['QuestionID']
            question.name = question_payload['DataExportTag']
            question.prompt = question_payload['QuestionText']
            question.type = question_payload['QuestionType']
            question.subtype = question_payload['SubSelector']
            question.has_carry_forward_prompts = True
            carry_forward_locator = question_payload['DynamicChoices']['Locator']
            carry_forward_match = re.match('q://(QID\d+).+', carry_forward_locator)
            question.carry_forward_question_id = carry_forward_match.group(1)
            matrix_question.add_question(question)
            matrix_question.id = question.id
        return matrix_question
        
class QSFResponsesParser(object):

    def parse(object, question, question_payload, question_element):
        question.subtype = question_payload['Selector']
        if question_payload.get('Choices') and len(question_payload['Choices']) > 0:
            question.response_order = question_payload['ChoiceOrder']
            for code, response in question_payload['Choices'].iteritems():
                question.add_response(response['Display'], code)
        
    def carry_forward_responses(self, questions):
        dynamic_questions = [question for question in questions if question.has_carry_forward_responses == True]
        for dynamic_question in dynamic_questions:
            matching_question = next((question for question in questions if question.id == dynamic_question.carry_forward_question_id), None)
            dynamic_question.response_order = matching_question.response_order
            for response in matching_question.responses:
                dynamic_question.add_response(response.response, response.code)
        return questions


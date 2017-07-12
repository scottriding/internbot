import re
from survey import Survey
from block import Blocks, Block
from question import Questions, Question, CompositeQuestion

class QSFSurveyParser(object):

    def parse(self, survey_element):
        survey = self.create_survey(survey_element)
        return survey
        
    def create_survey(self, survey_element):
        return Survey(self.parse_name(survey_element))
        
    def parse_name(self, survey_element):
        return survey_element['SurveyName']
        
class QSFBlockFlowParser(object):

    def __init__(self):
        self.__block_ids = []
    
    def parse(self, flow_element):
        flow_payload = flow_element['Payload']['Flow']
        for block in flow_payload:
            if block['Type'] != 'Branch' and block['Type'] != 'EmbeddedData' and block['Type'] != 'EndSurvey':
                self.__block_ids.append(block['ID'])
        return self.__block_ids
            
class QSFBlocksParser(object):

    def __init__(self):
        self.__blocks = Blocks()

    def parse(self, blocks_element):
        for block_element in blocks_element['Payload']:
            if block_element['Type'] != 'Trash':
                block = self.block_basics(block_element)
                self.__blocks.add(block)
        return self.__blocks
        
    def block_basics(self, block_element):
        block = Block(block_element['Description'])
        block.blockid = block_element['ID']
        self.assign_question_id(block_element, block)
        return block
        
    def assign_question_id(self, block_element, block):
        for question_id in block_element['BlockElements']:
            if question_id['Type'] == 'Question':
                block.assign_id(question_id['QuestionID'])

class QSFQuestionsParser(object):

    def __init__(self):
        self.matrix_parser = QSFQuestionsMatrixParser()
        self.response_parser = QSFResponsesParser()
        self.__questions = []

    def parse(self, question_elements):
        for question_element in question_elements:
            question_payload = question_element['Payload']
            question = self.parse_question_basics(question_payload)
            
            if question_payload.get('DynamicChoices') and question.type != 'Matrix':
                self.assign_carry_forward(question, question_payload)

            if question.type == 'Matrix':
                matrix_question = self.matrix_parser.parse(question_payload)
                self.__questions.append(matrix_question)
            else:
                self.response_parser.parse(question, question_payload, question_element)
                self.__questions.append(question)

        self.__questions = self.response_parser.carry_forward_responses(self.__questions)
        self.__questions = self.carry_forward_prompts(self.__questions)
        return self.__questions
        
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
                matrix_question.add(question)
                matrix_question.id = question_payload['QuestionID']
                matrix_question.question_order = question_payload['ChoiceOrder']
                matrix_question.name = question_payload['DataExportTag']
                matrix_question.subtype = question_payload['SubSelector']
                matrix_question.prompt = prompt['Display']
        else:
            question = Question()
            question.id = question_payload['QuestionID']
            question.name = question_payload['DataExportTag']
            question.prompt = question_payload['QuestionText']
            question.subtype = question_payload['SubSelector']
            question.has_carry_forward_prompts = True
            carry_forward_locator = question_payload['DynamicChoices']['Locator']
            carry_forward_match = re.match('q://(QID\d+).+', carry_forward_locator)
            question.carry_forward_question_id = carry_forward_match.group(1)
            matrix_question.add(question)
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


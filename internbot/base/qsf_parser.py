import json
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
        survey.update_questions(questions)
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
                        question = Question(question_id['QuestionID'])
                        block.add_question(question)
                blocks.add(block)
        return blocks

class QSFQuestionsParser(object):

    def __init__(self):
        self.matrix_parser = QSFQuestionsMatrixParser()
        
    def parse(self, question_elements):
        questions = []
        matrix_questions = []
        for question_element in question_elements:
            question_payload = question_element['Payload']
            question = Question(question_payload['QuestionID'])
            question.name = question_payload['DataExportTag']
            question.prompt = question_payload['QuestionText']
            question.type = question_payload['QuestionType']
            if not(question.type == 'Matrix'):
                question.subtype = question_payload['Selector']
                if question_payload.get('Choices') and len(question_payload['Choices']) > 0:
                    for code, response in question_payload['Choices'].iteritems():
                        question.add_response(response['Display'], code)
                questions.append(question)        
            elif question.type == 'Matrix' and (len(question_payload['Choices']) > 0):
                matrix_prompts = question_payload['Choices']
                matrix_questions = self.matrix_parser.parse(question_payload, matrix_prompts)   
                for question in matrix_questions:
                    questions.append(question)
        print questions
        return questions
        
class QSFQuestionsMatrixParser(object):

    def parse(self, question_payload, matrix_prompts):
        matrix_questions = []
        for code, prompt in matrix_prompts.iteritems():
            matrix_question = Question(str(question_payload['QuestionID']) + '_' + code)
            matrix_question.subtype = question_payload['SubSelector']
            matrix_question.name = str(question_payload['DataExportTag']) + '_' + code
            matrix_question.prompt = prompt['Display']
            for code, response in question_payload['Answers'].iteritems():
                matrix_question.add_response(response['Display'], code)
            matrix_questions.append(matrix_question)
        return matrix_questions    



import re
from survey import Survey
from block import Blocks, Block
from question import Questions, Question, CompositeQuestion
from HTMLParser import HTMLParser

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
            self.basic_flow_structure(block)
        return self.__block_ids
        
    def basic_flow_structure(self, block):
        if block['Type'] == 'Block' or block['Type'] == 'Standard':
            self.__block_ids.append(block['ID'])
        else:
            self.layered_flow_structure(block)
        
    def layered_flow_structure(self, block):
        for type, value in block.iteritems():
            if (type == 'Flow' and len(value) > 0) and \
               (type == 'Flow' and value[0].get('EmbeddedData') is None):
                self.grab_blockid_layer_details(value)

    def grab_blockid_layer_details(self, layer_details):
        for detail in layer_details:
            if detail.get('ID') is not None:
                self.__block_ids.append(detail['ID'])
            elif detail.get('Flow') is not None and \
                 detail['Flow'][0].get('ID') is not None:
                self.__block_ids.append(detail['Flow'][0]['ID']) 

class QSFBlocksParser(object):

    def __init__(self):
        self.__blocks = Blocks()

    def parse(self, blocks_element):
        payload = blocks_element['Payload']
        try:
            self.basic_block_structure(payload)
        except TypeError:
            self.layered_block_structure(payload)
        return self.__blocks
        
    def basic_block_structure(self, payload):
        for block_element in payload:
            if block_element['Type'] != 'Trash':
                block = self.block_details(block_element)
                self.__blocks.add(block)
    
    def layered_block_structure(self, payload):
        for key, value in payload.iteritems():
            if value['Type'] != 'Trash':
                block = self.block_details(value)
                self.__blocks.add(block)
    
    def block_details(self, block_element):
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
        self.carryforwardparser = QSFCarryForwardParser()
        self.hotspot_parser = QSFQuestionHotSpotParser()
        self.__questions = []

    def parse(self, question_elements):
        for question_element in question_elements:
            question_payload = question_element['Payload']
            question = self.question_details(question_payload)
            
            if question_payload.get('DynamicChoices') and \
               question.type != 'Matrix':
                self.carryforwardparser.assign_carry_forward(question, question_payload)

            if question.type == 'Matrix':
                matrix_question = self.matrix_parser.parse(question_payload)
                self.__questions.append(matrix_question)
            elif question.type == 'HotSpot':
                hotspot_question = self.hotspot_parser.parse(question, question_payload)
                self.__questions.append(hotspot_question)
            elif question.type == 'Meta':
                pass
            else:
                self.response_parser.parse(question, question_payload, question_element)
                self.__questions.append(question)

        self.__questions = self.carryforwardparser.carry_forward_responses(self.__questions)
        self.__questions = self.carryforwardparser.carry_forward_prompts(self.__questions)
        return self.__questions
        
    def question_details(self, question_payload):
        question = Question()
        question.id = question_payload['QuestionID']
        question.name = question_payload['DataExportTag']
        question.prompt = self.strip_tags( \
                          question_payload['QuestionText'].encode('ascii', 'ignore'))
        question.type = question_payload['QuestionType']
        return question

    def strip_tags(self, html):
        html_parser = MLStripper()
        html_parser.feed(html)
        return html_parser.get_data()
        

class QSFQuestionsMatrixParser(object):

    def parse(self, question_payload):
        matrix_question = CompositeQuestion()
        if question_payload.get('DynamicChoices') is None:
            self.basic_matrix(question_payload, matrix_question)       
        else:
            self.dynamic_matrix(question_payload, matrix_question)
        return matrix_question
       
    def dynamic_matrix(self, question_payload, matrix_question):
        responses = question_payload['Answers']
        self.matrix_details(matrix_question, question_payload)
        matrix_question.has_carry_forward_prompts = True
        carry_forward_locator = question_payload['DynamicChoices']['Locator']
        carry_forward_match = re.match('q://(QID\d+).+', carry_forward_locator)
        matrix_question.carry_forward_question_id = carry_forward_match.group(1)
        for code, response in responses.iteritems():
            response_name = self.strip_tags( \
                            response['Display'].encode('ascii','ignore'))
            matrix_question.add_response(response_name, code)
            
    def basic_matrix(self, question_payload, matrix_question):
        prompts = question_payload['Choices']
        responses = question_payload['Answers']
        for code, prompt in prompts.iteritems():
            question = self.question_details(code, prompt, question_payload, responses)
            matrix_question.add_question(question)
            matrix_question.question_order = question_payload['ChoiceOrder']
            self.matrix_details(matrix_question, question_payload)
    
    def question_details(self, code, prompt, question_payload, responses):
        question = Question()
        question.id = '%s_%s' % (str(question_payload['QuestionID']), code)
        question.code = code
        question.type = question_payload['QuestionType']
        question.subtype = question_payload['SubSelector']
        if question_payload['ChoiceDataExportTags']:
            question.name = question_payload['ChoiceDataExportTags'][code]    
        else:
            question.name = '%s_%s' % (str(question_payload['DataExportTag']), code)
        question.prompt = self.strip_tags(prompt['Display'].encode('ascii', 'ignore'))
        question.response_order = question_payload['AnswerOrder']
        for code, response in responses.iteritems():
            response_name = self.strip_tags(response['Display'] \
                            .encode('ascii','ignore'))
            question.add_response(response_name, code)
        return question
        
    def matrix_details(self, matrix_question, question_payload):
        matrix_question.id = question_payload['QuestionID']
        matrix_question.name = question_payload['DataExportTag']
        matrix_question.prompt = self.strip_tags( \
                                 question_payload['QuestionText'].encode('ascii', 'ignore'))
        matrix_question.subtype = question_payload['SubSelector']

    def strip_tags(self, html):
        html_parser = MLStripper()
        html_parser.feed(html)
        return html_parser.get_data()

class QSFQuestionHotSpotParser(object):

    def parse(self, question, question_payload):
        hotspot_question = CompositeQuestion()
        hotspot_question.name = question.name
        hotspot_question.prompt = question.prompt
        hotspot_question.id = question_payload['QuestionID']
        self.basic_hotspot(question_payload, hotspot_question)    
        return hotspot_question

    def basic_hotspot(self, question_payload, hotspot_question):
        if question_payload.get('Choices') and len(question_payload['Choices']) > 0:
            if question_payload.get('ChoiceOrder') and \
               len(question_payload['ChoiceOrder']) > 0: 
                hotspot_question.question_order = question_payload['ChoiceOrder']
            for code, question in question_payload['Choices'].iteritems():
                sub_question = Question()
                sub_question.id = '%s_%s' % (hotspot_question.id, code)
                sub_question.code = code
                sub_question.type = question_payload['QuestionType']
                sub_question.subtype = question_payload['Selector']
                sub_question.name = '%s_%s' % (hotspot_question.name, code)
                sub_question.prompt = question['Display']
                sub_question.add_response('0',0)
                sub_question.add_response('1',1)
                hotspot_question.add_question(sub_question)
        
class QSFResponsesParser(object):

    def parse(object, question, question_payload, question_element):
        question.subtype = question_payload['Selector']
        if question_payload.get('Choices') and len(question_payload['Choices']) > 0:
            if question_payload.get('ChoiceOrder') and \
               len(question_payload['ChoiceOrder']) > 0: 
                question.response_order = question_payload['ChoiceOrder']
            if question.subtype == 'NPS':
                for iteration in question_payload['Choices']:
                    for response, code in iteration.iteritems():
                        question.add_response(code, code)
            else:
                for code, response in question_payload['Choices'].iteritems():
                    question.add_response(response['Display'], code)
                                   

class QSFCarryForwardParser(object):

    def assign_carry_forward(self, question, question_payload):
        question.has_carry_forward_responses = True
        carry_forward_locator = question_payload['DynamicChoices']['Locator']
        carry_forward_match = re.match('q://(QID\d+).+', carry_forward_locator)
        question.carry_forward_question_id = carry_forward_match.group(1)

    def carry_forward_prompts(self, questions):
        dynamic_questions = [question for question in questions \
                            if question.has_carry_forward_prompts == True]
        carried_forward_questions = []
        for dynamic_question in dynamic_questions:
            matching_question = next((question for question in questions \
                                if question.id == dynamic_question.carry_forward_question_id), None)
            if matching_question.type == 'Composite':
                self.composite_prompts(matching_question, \
                                       dynamic_question,  \
                                       carried_forward_questions)
            else:
                self.basic_prompts(matching_question, \
                                   dynamic_question,  \
                                   carried_forward_questions)                
        return questions        

    def composite_prompts(self, matching_question, dynamic_question, carried_forward_questions):
        dynamic_question.question_order = matching_question.question_order
        for sub_question in matching_question.questions:
            question = Question()
            question.prompt = sub_question.prompt
            question.code = sub_question.code
            question.id = '%s_%s' % (dynamic_question.id, sub_question.code)
            question.name = '%s_%s' % (dynamic_question.name, sub_question.code)
            question.response_order = sub_question.response_order
            for response in sub_question.responses:
                question.add_response(response.response, response.code)
            dynamic_question.add_question(question)
            carried_forward_questions.append(question)

    def basic_prompts(self, matching_question, dynamic_question, carried_forward_questions):
        dynamic_question.question_order = matching_question.response_order
        for response in matching_question.responses:
            question = Question()
            question.prompt = response.response
            question.code = response.code
            question.id = '%s_%s' % (dynamic_question.id, response.code)
            question.name = '%s_%s' % (dynamic_question.name, response.code)
            question.response_order = matching_question.response_order
            for response in dynamic_question.temp_responses:
                question.add_response(response.response, response.code)
            dynamic_question.add_question(question)
            carried_forward_questions.append(question)

    def carry_forward_responses(self, questions):
        dynamic_questions = [question for question in questions \
                            if question.has_carry_forward_responses == True]
        for dynamic_question in dynamic_questions:
            matching_question = next((question for question in questions \
                                    if question.id == dynamic_question.carry_forward_question_id), None)
            dynamic_question.response_order = matching_question.response_order
            for response in matching_question.responses:
                dynamic_question.add_response(response.response, response.code)
        return questions

class MLStripper(HTMLParser):
    def __init__(self):
        self.reset()
        self.fed = []
    
    def handle_data(self, data):
        self.fed.append(data)
    
    def get_data(self):
        return ' '.join(self.fed)
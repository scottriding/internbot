import re
from survey import Survey
from block import Blocks, Block
from scoring import Scorings
from question import Questions, Question 
from question import CompositeQuestion, CompositeMatrix, CompositeMultipleSelect
from question import CompositeHotSpot, CompositeConstantSum
from HTMLParser import HTMLParser

class QSFSurveyParser(object):

    def parse(self, survey_element):
        survey = self.create_survey(survey_element)
        return survey
        
    def create_survey(self, survey_element):
        return Survey(self.parse_name(survey_element))
        
    def parse_name(self, survey_element):
        return survey_element['SurveyName']

class QSFScoringParser(object):

    def __init__(self):
        self.__scoring_ids = Scorings()

    def parse(self, sco_element):
        sco_payload = sco_element[0]['Payload']['ScoringCategories']
        for score in sco_payload:
            self.__scoring_ids.add(score["ID"], score["Name"], score["Description"])
        return self.__scoring_ids
 
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
        self.hotspot_parser = QSFQuestionHotSpotParser()
        self.multiple_parser = QSFMultipleSelectParser()
        self.constant_parser = QSFConstantSumParser()
        self.carryforwardparser = QSFCarryForwardParser()
        self.response_parser = QSFResponsesParser()
        self.__questions = []

    def parse(self, question_elements):
        for question_element in question_elements:
            question_payload = question_element['Payload']
            question = self.question_details(question_payload)
            self.parse_dynamic(question, question_payload)
            self.parse_type(question, question_payload, question_element)
        self.__questions = self.carryforwardparser.carry_forward(self.__questions)
        return self.__questions

    def parse_dynamic(self, question, question_payload):
        if question_payload.get('DynamicChoices'):
            self.carryforwardparser.assign_carry_forward(question, question_payload)
            if question_payload.get('Choices') is True:
                if len(question_payload['Choices']) > 0:
                    question.has_mixed_responses = True

    def parse_type(self, question, question_payload, question_element):
        if question.type == 'Matrix':
            if question_payload['Selector'] == 'MaxDiff':
                pass
            else:
                matrix_question = self.matrix_parser.parse(question_payload)
                self.__questions.append(matrix_question)
        elif question.type == 'HotSpot':
            hotspot_question = self.hotspot_parser.parse(question, question_payload)
            self.__questions.append(hotspot_question)
        elif question.type == 'MC' and question.subtype == 'MAVR':
            multiple_select = self.multiple_parser.parse(question, question_payload)
            self.__questions.append(multiple_select)
        elif question.type == 'MC' and question.subtype == 'MACOL':
            multiple_select = self.multiple_parser.parse(question, question_payload)
            self.__questions.append(multiple_select)
        elif question.type == 'TE':
            if question_payload['Validation'].get('Settings') is not None:
                if question_payload['Validation']['Settings'].get('ContentType') is not None:
                    question.subtype = question_payload['Validation']['Settings']['ContentType']
            self.__questions.append(question)
        elif question.type == 'CS':
            constant_sum = self.constant_parser.parse(question, question_payload)
            self.__questions.append(constant_sum)
        elif question.type == 'Meta':
            pass
        else:
            self.response_parser.parse(question, question_payload, question_element)
            self.__questions.append(question)
        
    def question_details(self, question_payload):
        question = Question()
        question.id = question_payload['QuestionID']
        question.name = question_payload['DataExportTag']
        question.prompt = self.strip_tags( \
                          question_payload['QuestionText'].encode('ascii', 'ignore'))
        question.type = question_payload['QuestionType']
        if question_payload.get('Selector') is not None:
            question.subtype = question_payload['Selector']
        return question

    def strip_tags(self, html):
        html_parser = MLStripper()
        html_parser.feed(html)
        return html_parser.get_data()

class QSFQuestionsMatrixParser(object):

    def __init__(self):
        self.carry_forward = QSFCarryForwardParser()

    def parse(self, question_payload):
        matrix_question = CompositeMatrix()
        if question_payload.get('DynamicChoices') is None:
            self.basic_matrix(question_payload, matrix_question)      
        elif question_payload.get('DynamicChoices') and \
             len(question_payload['Choices']) > 0:
            self.mixed_matrix(question_payload, matrix_question)
        else:
            self.dynamic_matrix(question_payload, matrix_question)          
        return matrix_question
       
    def dynamic_matrix(self, question_payload, matrix_question):
        responses = question_payload['Answers']
        self.matrix_details(matrix_question, question_payload)
        self.carry_forward.assign_carry_forward(matrix_question, question_payload)
        for code, response in responses.iteritems():
            response_name = self.strip_tags( \
                            response['Display'].encode('ascii','ignore'))
            matrix_question.add_response(response_name, code)

    def mixed_matrix(self, question_payload, matrix_question):
        self.basic_matrix(question_payload, matrix_question)
        self.dynamic_matrix(question_payload, matrix_question)
            
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
        if question_payload.get('SubSelector') is None:
            pass
        else:
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
        if question_payload.get('SubSelector') is None:
            pass
        else:
            matrix_question.subtype = question_payload['SubSelector']

    def strip_tags(self, html):
        html_parser = MLStripper()
        html_parser.feed(html)
        return html_parser.get_data()

class QSFQuestionHotSpotParser(object):

    def parse(self, question, question_payload):
        hotspot_question = self.hotspot_details(question, question_payload)
        self.basic_hotspot(question_payload, hotspot_question)    
        return hotspot_question

    def basic_hotspot(self, question_payload, hotspot_question):
        if question_payload.get('Choices') and len(question_payload['Choices']) > 0:
            if question_payload.get('ChoiceOrder') and \
               len(question_payload['ChoiceOrder']) > 0: 
                hotspot_question.question_order = question_payload['ChoiceOrder']
            self.question_details(hotspot_question, question_payload)

    def hotspot_details(self, question, question_payload):
        hotspot = CompositeHotSpot()
        hotspot.name = question.name
        hotspot.prompt = question.prompt
        hotspot.subtype = question.subtype
        hotspot.id = question_payload['QuestionID']
        return hotspot

    def question_details(self, hotspot, question_payload):
        for code, question in question_payload['Choices'].iteritems():
            sub_question = Question()
            sub_question.id = '%s_%s' % (hotspot.id, code)
            sub_question.code = code
            sub_question.type = question_payload['QuestionType']
            sub_question.subtype = question_payload['Selector']
            sub_question.name = '%s_%s' % (hotspot.name, code)
            sub_question.prompt = question['Display']
            sub_question.add_response('0',1)
            sub_question.add_response('1',2)
            hotspot.add_question(sub_question)

class QSFMultipleSelectParser(object):

    def __init__(self):
        self.carryforward = QSFCarryForwardParser()

    def parse(self, question, question_payload):
        multiple_select = self.multiselect_details(question_payload, question)
        if question.has_carry_forward_responses is False:
            self.basic_multiple(multiple_select, question_payload)     
        elif question.has_carry_forward_responses is True and \
             len(question_payload['Choices']) > 0:
            self.mixed_multiple(multiple_select, question_payload)
        else:
            self.dynamic_multiple(multiple_select, question_payload)  
        return multiple_select

    def dynamic_multiple(self, multiple_question, question_payload):
        self.carryforward.assign_carry_forward(multiple_question, question_payload)

    def mixed_multiple(self, multiple_question, question_payload):
        self.dynamic_multiple(multiple_question, question_payload)
        self.basic_multiple(multiple_question, question_payload)

    def basic_multiple(self, multiple_select, question_payload):
        if question_payload.get('Choices') and len(question_payload['Choices']) > 0:
            if question_payload.get('ChoiceOrder') and \
               len(question_payload['ChoiceOrder']) > 0: 
                multiple_select.question_order = question_payload['ChoiceOrder']
            self.question_details(question_payload, multiple_select)

    def multiselect_details(self, question_payload, question):
        multiselect = CompositeMultipleSelect()
        multiselect.name = question.name
        multiselect.prompt = question.prompt
        multiselect.subtype = question.subtype
        multiselect.id = question_payload['QuestionID']
        return multiselect

    def question_details(self, question_payload, multiple_select):
        for code, question in question_payload['Choices'].iteritems():
            sub_question = Question()
            sub_question.id = '%s_%s' % (multiple_select.id, code)
            sub_question.code = code
            sub_question.type = question_payload['QuestionType']
            sub_question.subtype = question_payload['Selector']
            sub_question.name = '%s_%s' % (multiple_select.name, code)
            sub_question.prompt = question['Display'].encode('ascii', 'ignore')
            sub_question.add_response('1',1)
            multiple_select.add_question(sub_question)

class QSFConstantSumParser(object):

    def __init__(self):
        self.carryforward = QSFCarryForwardParser()

    def parse(self, question, question_payload):
        constant_sum = self.constantsum_details(question_payload, question)
        if question.has_carry_forward_responses is False:
            self.basic_constant(constant_sum, question_payload)       
        elif question.has_carry_forward_responses is True and \
             question_payload['Choices'] > 0:
            self.mixed_constant(constant_sum, question_payload)
        else:
            self.dynamic_constant(constant_sum, question_payload)
        return constant_sum

    def dynamic_constant(self, constant_sum, question_payload):
        self.carryforward.assign_carry_forward(constant_sum, question_payload)    

    def mixed_constant(self, constant_sum, question_payload):
        self.dynamic_constant(constant_sum, question_payload)
        self.basic_constant(constant_sum, question_payload)

    def basic_constant(self, constant_sum, question_payload):
        if question_payload.get('Choices') and len(question_payload['Choices']) > 0:
            if question_payload.get('ChoiceOrder') and \
               len(question_payload['ChoiceOrder']) > 0: 
                constant_sum.question_order = question_payload['ChoiceOrder']
            self.question_details(question_payload, constant_sum)

    def constantsum_details(self, question_payload, question):
        constantsum = CompositeConstantSum()
        constantsum.name = question.name
        constantsum.prompt = question.prompt
        constantsum.subtype = question.subtype
        constantsum.id = question_payload['QuestionID']
        return constantsum

    def question_details(self, question_payload, constant_sum):
        for code, question in question_payload['Choices'].iteritems():
            sub_question = Question()
            sub_question.id = '%s_%s' % (constant_sum.id, code)
            sub_question.code = code
            sub_question.type = question_payload['QuestionType']
            sub_question.subtype = question_payload['Selector']
            sub_question.name = '%s_%s' % (constant_sum.name, code)
            sub_question.prompt = question['Display']
            sub_question.add_response(sub_question.prompt, sub_question.code)
            constant_sum.add_question(sub_question)

class QSFResponsesParser(object):

    def parse(self, question, question_payload, question_element):
        if question.has_mixed_responses is True:
            self.parse_mixed_responses(question, question_payload, question_element)
        elif question_payload.get('Choices') and len(question_payload['Choices']) > 0:
            self.parse_response_order(question, question_payload)
            if question.subtype == 'NPS':
                self.parse_NPS(question, question_payload)
            else:
                self.parse_basic(question, question_payload)

    def parse_mixed_responses(self, question, question_payload, question_element):
        pass
                
    def parse_response_order(self, question, question_payload):
        if question_payload.get('ChoiceOrder') and \
           len(question_payload['ChoiceOrder']) > 0: 
            question.response_order = question_payload['ChoiceOrder']

    def parse_NPS(self, question, question_payload):
        for iteration in question_payload['Choices']:
            for response, code in iteration.iteritems():
                question.add_response(code, code)

    def parse_basic(self, question, question_payload):
        for code, response in question_payload['Choices'].iteritems():
                question.add_response(response['Display'].encode('ascii','ignore'), code)
                 
class QSFCarryForwardParser(object):

    def assign_carry_forward(self, question, question_payload):
        question.has_carry_forward_responses = True
        carry_forward_locator = question_payload['DynamicChoices']['Locator']
        carry_forward_match = re.match('q://(QID\d+).+', carry_forward_locator)
        question.carry_forward_question_id = carry_forward_match.group(1)

    def carry_forward(self, questions):
        dynamic_questions = [question for question in questions \
                            if question.has_carry_forward_responses == True]
        for dynamic_question in dynamic_questions:
            if dynamic_question.type == 'CompositeMatrix':
                self.matrix_match(dynamic_question, questions)
            elif dynamic_question.type == 'CompositeMultipleSelect':
                self.multiselect_match(dynamic_question, questions)
            elif dynamic_question.type == 'MC':
                self.singleMulti_match(dynamic_question, questions)
        return questions

    def matrix_match(self, dynamic_question, questions):
        matching_question = next((question for question in questions \
                            if question.id == dynamic_question.carry_forward_question_id), None)
        if matching_question.type == 'CompositeMatrix':
            self.matrix_into_matrix(dynamic_question, matching_question)
        elif matching_question.type == 'CompositeMultipleSelect':
            self.multiselect_into_matrix(dynamic_question, matching_question)
        elif matching_question.type == 'MC':
            self.singleMulti_into_matrix(dynamic_question, matching_question)

    def multiselect_match(self, dynamic_question, questions):
        matching_question = next((question for question in questions \
                            if question.id == dynamic_question.carry_forward_question_id), None)
        if matching_question.type == 'CompositeMatrix':
            self.matrix_into_multiselect(dynamic_question, matching_question)
        elif matching_question.type == 'CompositeMultipleSelect':
            self.multiselect_into_multiselect(dynamic_question, matching_question)
        elif matching_question.type == 'MC':
            self.singleMulti_into_multiselect(dynamic_question, matching_question)

    def singleMulti_match(self, dynamic_question, questions):
        matching_question = next((question for question in questions \
                             if question.id == dynamic_question.carry_forward_question_id), None)
        if matching_question.type == 'CompositeMatrix':
            self.matrix_into_singleMulti(dynamic_question, matching_question)
        elif matching_question.type == 'CompositeMultipleSelect':
            self.multiselect_into_singleMulti(dynamic_question, matching_question)
        elif matching_question.type == 'MC':
            self.singleMulti_into_singleMulti(dynamic_question, matching_question)

    def matrix_into_matrix(self, dynamic_matrix, matching_matrix):
        dynamic_matrix.question_order = matching_matrix.question_order
        for question in matching_matrix.questions:
            sub_question = Question()
            sub_question.name ='%s_%s' % (dynamic_matrix.name, question.code)
            sub_question.id = '%s_%s' % (dynamic_matrix.name, question.code)
            sub_question.code = question.code
            sub_question.prompt = question.prompt
            for response in question.responses:
                sub_question.add_dynamic_response(response.response, response.code)
            dynamic_matrix.add_question(sub_question)

    def multiselect_into_matrix(self, dynamic_matrix, matching_multiselect):
        dynamic_matrix.question_order = matching_multiselect.question_order
        for question in matching_multiselect.questions:
            sub_question = Question()
            sub_question.name = '%s_%s' % (dynamic_matrix.name, question.code)
            sub_question.id = '%s_%s' % (dynamic_matrix.id, question.code)
            sub_question.code = question.code
            sub_question.prompt = question.prompt
            for sub_responses in dynamic_matrix.temp_responses:
                sub_question.add_dynamic_response(sub_responses.response, sub_responses.code)
            dynamic_matrix.add_question(sub_question)

    def singleMulti_into_matrix(self, dynamic_matrix, matching_MC):
        dynamic_matrix.question_order = matching_MC.response_order
        for response in matching_MC.responses:
            sub_question = Question()
            sub_question.name = '%s_%s' % (dynamic_matrix.name, response.code)
            sub_question.id = '%s_%s' % (dynamic_matrix.id, response.code)
            sub_question.code = response.code
            sub_question.prompt = response.response
            for sub_response in dynamic_matrix.temp_responses:
                sub_question.add_dynamic_response(sub_response.response, sub_response.code)
            dynamic_matrix.add_question(sub_question)

    def matrix_into_multiselect(self, multiselect, matrix):
        multiselect.question_order = matrix.question_order
        for question in matrix.questions:
            sub_question = Question()
            sub_question.name = '%s_x%s' % (matrix.name, question.code)
            sub_question.id = '%s_%s' % (multiselect.id, question.code)
            sub_question.code = question.code
            sub_question.prompt = question.prompt
            sub_question.add_dynamic_response('1', 1)
            multiselect.add_question(sub_question)

    def multiselect_into_multiselect(self, dynamic_multi, matching_multi):
        dynamic_multi.question_order = matching_multi.question_order
        for sub_question in matching_multi.questions:
            question = Question()
            question.name = '%s_x%s' % (dynamic_multi.name, sub_question.code)
            question.id = '%s_%s' % (dynamic_multi.id, sub_question.code)
            question.code = sub_question.code
            question.prompt = sub_question.prompt
            for response in sub_question.responses:
                question.add_dynamic_response(response.response, response.code)
            dynamic_multi.add_question(question)

    def singleMulti_into_multiselect(self, multiselect_question, matching_MC):
        multiselect_question.question_order = matching_MC.response_order
        for response in matching_MC.responses:
            sub_question = Question()
            sub_question.name = '%s_x%s' % (multiselect_question.name, response.code)
            sub_question.id = '%s_%s' % (multiselect_question.id, response.code)
            sub_question.code = response.code
            sub_question.prompt = response.response
            sub_question.add_dynamic_response('1', 1)
            multiselect_question.add_question(sub_question)

    def matrix_into_singleMulti(self, dynamic_MC, matching_matrix):
        dynamic_MC.response_order = matching_matrix.question_order
        for question in matching_matrix.questions:
            dynamic_MC.add_dynamic_response(question.prompt, question.code)

    def multiselect_into_singleMulti(self, dynamic_MC, matching_multiselect):
        dynamic_MC.response_order = matching_multiselect.question_order
        for sub_question in matching_multiselect.questions:
            dynamic_MC.add_dynamic_response(sub_question.prompt, sub_question.code)

    def singleMulti_into_singleMulti(self, dynamic_MC, matching_MC):
        dynamic_MC.response_order = matching_MC.response_order
        for response in matching_MC.responses:
            dynamic_MC.add_dynamic_response(response.response, response.code)

class MLStripper(HTMLParser):
    def __init__(self):
        self.reset()
        self.fed = []
    
    def handle_data(self, data):
        self.fed.append(data)
    
    def get_data(self):
        return ' '.join(self.fed)
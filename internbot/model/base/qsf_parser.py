from model.base import survey
from model.base import block
from model.base import scoring
from model.base import question

import re
from html import parser

class QSFSurveyParser(object):

    def parse(self, survey_element):
        survey = self.create_survey(survey_element)
        return survey
        
    def create_survey(self, survey_element):
        return survey.Survey(self.parse_name(survey_element))
        
    def parse_name(self, survey_element):
        return survey_element['SurveyName']

class QSFScoringParser(object):

    def __init__(self):
        self.__scoring_ids = scoring.Scorings()

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
        for type, value in block.items():
            if type == 'Flow' and len(value) > 0:
                self.grab_blockid_layer_details(value)

    def grab_blockid_layer_details(self, layer_details):
        for detail in layer_details:
            if detail.get('ID') is not None:
                self.__block_ids.append(detail['ID'])
            elif detail.get('Flow') is not None:
                self.layered_flow_structure(detail)

class QSFBlocksParser(object):

    def __init__(self):
        self.__blocks = block.Blocks()

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
        for key, value in payload.items():
            if value['Type'] != 'Trash':
                block = self.block_details(value)
                self.__blocks.add(block)
    
    def block_details(self, block_element):
        new_block = block.Block(block_element['Description'])
        new_block.blockid = block_element['ID']
        self.assign_question_id(block_element, new_block)
        return new_block
        
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
        elif question.type == 'SBS':
            print("\n"+question.name+": Side by side question type is not yet implemented.\n"
                                     "             ***This question will need to be input manually.***\n"
                                     "    ***You may see more error messages about this question below***\n")
        else:
            self.response_parser.parse(question, question_payload, question_element)
            self.__questions.append(question)
       
    def question_details(self, question_payload):
        new_question = question.Question()
        new_question.id = question_payload['QuestionID']
        new_question.name = question_payload['DataExportTag']
        new_question.prompt = self.strip_tags(question_payload['QuestionText'])
        new_question.type = question_payload['QuestionType']
        if question_payload.get('Selector') is not None:
            new_question.subtype = question_payload['Selector']
        return new_question

    def strip_tags(self, html):
        html_stripper = MLStripper()
        html_stripper.feed(html)
        return html_stripper.get_data()

class QSFQuestionsMatrixParser(object):

    def __init__(self):
        self.carry_forward = QSFCarryForwardParser()

    def parse(self, question_payload):
        matrix_question = question.CompositeMatrix()
        self.matrix_details(matrix_question, question_payload)

        dynamic_statements = False
        mixed_statements = False
        dynamic_answers = False
        mixed_answers = False

        if question_payload.get('DynamicAnswers') is not None:
            if len(question_payload['Answers']) > 0:
                mixed_answers = True
            else:
                dynamic_answers = True

        if question_payload.get('DynamicChoices') is not None:
            if len(question_payload['Choices']) > 0:
                mixed_statements = True
            else:
                dynamic_statements = True

        #### figure out what kind of carry_forward madness is happening here

        # both statements and answers are a mix of carry_forward and additional content
        if mixed_answers == True and mixed_statements == True:
            matrix_question = self.both_mixed_matrix(question_payload, matrix_question)
        # both statements and answers are ONLY carry_forward
        elif dynamic_statements and dynamic_answers == True:
            matrix_question = self.both_dynamic_matrix(question_payload, matrix_question)
        # JUST carry_forward statements and a mix in the answers
        elif dynamic_statements and mixed_answers == True:
            matrix_question = self.dynamic_statements_mixed_answers_matrix(question_payload, matrix_question)
        # A mix of old and new statements and carry_forward answers
        elif mixed_statements == True and dynamic_answers == True:
            matrix_question = self.mixed_statements_dynamic_answers_matrix(question_payload, matrix_question)
        
        #### at this point just check statements
        # carry_forward statements ONLY
        elif dynamic_statements == True:
            matrix_question = self.dynamic_statements_matrix(question_payload, matrix_question)
        # mixed statements ONLY
        elif mixed_statements == True:
            matrix_question = self.mixed_statements_matrix(question_payload, matrix_question)

        #### at this point just check answers
        # carry_forward answers ONLY
        elif dynamic_answers == True:
            matrix_question = self.dynamic_answers_matrix(question_payload, matrix_question)
        # mixed answers ONLY
        elif mixed_answers == True:
            matrix_question = self.mixed_answers_matrix(question_payload, matrix_question)

        #### all edge cases have been checked -- this should be a basic matrix
        else:
            matrix_question = self.basic_matrix(question_payload, matrix_question)
      
        return matrix_question
       
    def both_mixed_matrix(self, question_payload, matrix_question):
        matrix_question = self.basic_matrix(question_payload, matrix_question)
        matrix_question = self.both_dynamic_matrix(question_payload, matrix_question)
        return matrix_question

    def dynamic_statements_mixed_answers_matrix(self, question_payload, matrix_question):
        self.carry_forward.assign_carry_forward_statements(matrix_question, question_payload)
        self.carry_forward.assign_carry_forward_answers(matrix_question, question_payload)
        responses = question_payload['Answers']
        for code, response in responses.items():
            response_name = self.strip_tags(response['Display'].encode('ascii','ignore'))
            matrix_question.add_response(response_name, code)
        return matrix_question

    def mixed_statements_dynamic_answers_matrix(self, question_payload, matrix_question):
        self.carry_forward.assign_carry_forward_answers(matrix_question, question_payload)
        self.carry_forward.assign_carry_forward_statements(matrix_question, question_payload)
        prompts = question_payload['Choices']
        for code, prompt in prompts.items():
            question = self.question_details(code, prompt, question_payload, prompts)
            matrix_question.add_question(question)
            matrix_question.question_order = question_payload['ChoiceOrder']
        return matrix_question

    def both_dynamic_matrix(self, question_payload, matrix_question):
        self.carry_forward.assign_carry_forward_answers(matrix_question, question_payload)
        self.carry_forward.assign_carry_forward_statements(matrix_question, question_payload)
        return matrix_question

    def dynamic_statements_matrix(self, question_payload, matrix_question):
        responses = question_payload['Answers']
        self.carry_forward.assign_carry_forward_statements(matrix_question, question_payload)
        for code, response in responses.items():
            response_name = self.strip_tags(response['Display'].encode('ascii','ignore'))
            matrix_question.add_response(response_name, code)
        return matrix_question

    def dynamic_answers_matrix(self, question_payload, matrix_question):
        prompts = question_payload['Choices']
        self.carry_forward.assign_carry_forward_answers(matrix_question, question_payload)
        for code, prompt in prompts.items():
            question = self.question_details(code, prompt, question_payload, prompts)
            matrix_question.add_question(question)
            matrix_question.question_order = question_payload['ChoiceOrder']
        return matrix_question

    def mixed_statements_matrix(self, question_payload, matrix_question):
        matrix_question = self.basic_matrix(question_payload, matrix_question)
        matrix_question = self.dynamic_statements_matrix(question_payload, matrix_question)
        return matrix_question

    def mixed_answers_matrix(self, question_payload, matrix_question):
        matrix_question = self.basic_matrix(question_payload, matrix_question)
        matrix_question = self.dynamic_answers_matrix(question_payload, matrix_question)
        return matrix_question
            
    def basic_matrix(self, question_payload, matrix_question):
        prompts = question_payload['Choices']
        responses = question_payload['Answers']
        for code, prompt in prompts.items():
            question = self.question_details(code, prompt, question_payload, responses)
            matrix_question.add_question(question)
            matrix_question.question_order = question_payload['ChoiceOrder']
        self.matrix_details(matrix_question, question_payload)
        return matrix_question
    
    def question_details(self, code, prompt, question_payload, responses):
        new_question = question.Question()
        new_question.id = '%s_%s' % (str(question_payload['QuestionID']), code)
        new_question.code = code
        new_question.type = question_payload['QuestionType']
        if question_payload.get('SubSelector') is None:
            pass
        else:
            new_question.subtype = question_payload['SubSelector']
        if question_payload['ChoiceDataExportTags']:
            new_question.name = question_payload['ChoiceDataExportTags'][code]    
        else:
            new_question.name = '%s_%s' % (str(question_payload['DataExportTag']), code)
        new_question.prompt = self.strip_tags(prompt['Display'].encode('ascii', 'ignore'))
        new_question.response_order = question_payload['AnswerOrder']

        for code, response in responses.items():
            try:
                response_name = self.strip_tags(response['Display'].encode('ascii','ignore'))
                new_question.add_response(response_name, code)
            except:
                pass
        return new_question
        
    def matrix_details(self, matrix_question, question_payload):
        matrix_question.id = question_payload['QuestionID']
        matrix_question.name = question_payload['DataExportTag']
        matrix_question.prompt = self.strip_tags(question_payload['QuestionText'].encode('ascii', 'ignore'))
        if question_payload.get('SubSelector') is None:
            pass
        else:
            matrix_question.subtype = question_payload['SubSelector']
        return matrix_question

    def strip_tags(self, html):
        html_stripper = MLStripper()
        html_stripper.feed(html.decode("utf-8"))
        return html_stripper.get_data()

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

    def hotspot_details(self, current_question, question_payload):
        hotspot = question.CompositeHotSpot()
        hotspot.name = current_question.name
        hotspot.prompt = current_question.prompt
        hotspot.subtype = current_question.subtype
        hotspot.id = question_payload['QuestionID']
        return hotspot

    def question_details(self, hotspot, question_payload):
        for code, current_question in question_payload['Choices'].items():
            sub_question = question.Question()
            sub_question.id = '%s_%s' % (hotspot.id, code)
            sub_question.code = code
            sub_question.type = question_payload['QuestionType']
            sub_question.subtype = question_payload['Selector']
            sub_question.name = '%s_%s' % (hotspot.name, code)
            sub_question.prompt = current_question['Display']
            sub_question.add_response('0','1')
            sub_question.add_response('1','2')
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

    def multiselect_details(self, question_payload, current_question):
        multiselect = question.CompositeMultipleSelect()
        multiselect.name = current_question.name
        multiselect.prompt = current_question.prompt
        multiselect.subtype = current_question.subtype
        multiselect.id = question_payload['QuestionID']
        return multiselect

    def question_details(self, question_payload, multiple_select):
        for code, parent_question in question_payload['Choices'].items():
            sub_question = question.Question()
            sub_question.id = '%s_%s' % (multiple_select.id, code)
            sub_question.code = code
            sub_question.type = question_payload['QuestionType']
            sub_question.subtype = question_payload['Selector']
            sub_question.name = '%s_%s' % (multiple_select.name, code)
            sub_question.prompt = self.convert_prompt_from_byte_str(parent_question['Display'].encode('ascii', 'ignore'))
            sub_question.add_response('1',1)
            multiple_select.add_question(sub_question)

    def convert_prompt_from_byte_str(self, prompt):
        prompt = str(prompt)
        if len(prompt):
            if prompt[0] is "b" and (prompt[len(prompt) - 1] is "'" or prompt[len(prompt) - 1] is '"'):
                converted = prompt[2: len(prompt) - 1]
                return converted
        return prompt

class QSFConstantSumParser(object):

    def __init__(self):
        self.carryforward = QSFCarryForwardParser()

    def parse(self, question, question_payload):
        constant_sum = self.constantsum_details(question_payload, question)
        if question.has_carry_forward_responses is False:
            self.basic_constant(constant_sum, question_payload)       
        elif question.has_carry_forward_responses is True and \
             len(question_payload['Choices']) > 0:
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

    def constantsum_details(self, question_payload, current_question):
        constantsum = question.CompositeConstantSum()
        constantsum.name = current_question.name
        constantsum.prompt = current_question.prompt
        constantsum.subtype = current_question.subtype
        constantsum.id = question_payload['QuestionID']
        return constantsum

    def question_details(self, question_payload, constant_sum):
        for code, parent_question in question_payload['Choices'].items():
            sub_question = question.Question()
            sub_question.id = '%s_%s' % (constant_sum.id, code)
            sub_question.code = code
            sub_question.type = question_payload['QuestionType']
            sub_question.subtype = question_payload['Selector']
            sub_question.name = '%s_%s' % (constant_sum.name, code)
            sub_question.prompt = parent_question['Display']
            sub_question.add_response(sub_question.prompt, sub_question.code)
            constant_sum.add_question(sub_question)

class QSFResponsesParser(object):

    def __init__(self):
        self.__response_order = {}

    def parse(self, question, question_payload, question_element):
        if question.has_mixed_responses is True:
            self.parse_mixed_responses(question, question_payload, question_element)
        elif question_payload.get('Choices') and len(question_payload['Choices']) > 0:
            if question.subtype == 'NPS':
                self.parse_NPS(question, question_payload)
            else:
                self.parse_response_order(question, question_payload)
                self.parse_basic(question, question_payload)
        self.__response_order.clear()

    def parse_mixed_responses(self, question, question_payload, question_element):
        pass
                
    def parse_response_order(self, question, question_payload):
        if question_payload.get('ChoiceOrder') and \
           len(question_payload['ChoiceOrder']) > 0:
            question.response_order = question_payload['ChoiceOrder']

        if question_payload.get('RecodeValues') is not None:
            response_values = []
            for old_value, new_value in question_payload['RecodeValues'].items():
                response_values.append(new_value)
                if question.response_order.__contains__(old_value):
                    question.response_order.remove(old_value)
                self.__response_order[old_value] = new_value
            question.response_order = response_values

    def parse_NPS(self, question, question_payload):
        if question_payload.get('ChoiceOrder') and len(question_payload['ChoiceOrder']) > 0:
            if question_payload.get('RecodeValues') is not None:
                question.response_order = question_payload['RecodeValues']
            else:
                question.response_order = question_payload['ChoiceOrder']

        for iteration in question_payload['Choices']:
            for response, value in iteration.items():
                question.add_response(value, value)

        for old_value, new_value in self.__response_order.items():
            matching_response = next((response for response in question.responses if response.value == old_value), None)
            if matching_response is not None:
                matching_response.value = new_value
                matching_response.label = self.convert_response_from_byte_str(matching_response)

    def parse_basic(self, question, question_payload):
        for value, label in question_payload['Choices'].items():
            question.add_response(label['Display'], value)
        for old_value, new_value in self.__response_order.items():
            matching_response = next((response for response in question.responses if response.value == old_value), None)
            if matching_response is not None:
                matching_response.value = new_value
                matching_response.label = self.convert_response_from_byte_str(matching_response.label)

    def convert_response_from_byte_str(self, response):
        if len(response) > 0:
            if response[0] is "b" and (response[len(response) - 1] is "'" or response[len(response) - 1] is '"'):
                converted = response[2: len(response) - 1]
                return converted
        return response
                 
class QSFCarryForwardParser(object):

    def assign_carry_forward_answers(self, question, question_payload):
        question.has_carry_forward_answers = True
        carry_forward_locator = question_payload['DynamicAnswers']['Locator']
        carry_forward_match = re.match('q://(QID\d+).+', carry_forward_locator)
        question.carry_forward_answers_id = carry_forward_match.group(1)

    def assign_carry_forward_statements(self, question, question_payload):
        question.has_carry_forward_statements = True
        carry_forward_locator = question_payload['DynamicChoices']['Locator']
        carry_forward_match = re.match('q://(QID\d+).+', carry_forward_locator)
        question.carry_forward_statements_id = carry_forward_match.group(1)

    def assign_carry_forward(self, question, question_payload):
        question.has_carry_forward_responses = True
        carry_forward_locator = question_payload['DynamicChoices']['Locator']
        carry_forward_match = re.match('q://(QID\d+).+', carry_forward_locator)
        question.carry_forward_question_id = carry_forward_match.group(1)

    def carry_forward(self, questions):
        dynamic_questions = [question for question in questions \
                            if question.has_carry_forward_responses == True]

        statement_questions = [question for question in questions \
                            if question.has_carry_forward_statements == True]

        answer_questions = [question for question in questions \
                            if question.has_carry_forward_answers == True]

        for dynamic_question in dynamic_questions:
            if dynamic_question.type == 'CompositeMultipleSelect':
                self.multiselect_match(dynamic_question, questions)
            elif dynamic_question.type == 'MC':
                self.singleMulti_match(dynamic_question, questions)
            elif dynamic_question.type == "CompositeConstantSum":
                self.constantsum_match(dynamic_question, questions)

        for statement_question in statement_questions:
            self.statement_match(statement_question, questions)

        for answer_question in answer_questions:
            self.answer_match(answer_question, questions)

        return questions

    def statement_match(self, dynamic_question, questions):
        matching_question = next((question for question in questions \
                            if question.id == dynamic_question.carry_forward_statements_id), None)
        if matching_question.type == 'CompositeMatrix':
            self.matrix_into_matrix(dynamic_question, matching_question)
        elif matching_question.type == 'CompositeMultipleSelect':
            self.multiselect_into_matrix(dynamic_question, matching_question)
        elif matching_question.type == 'MC':
            self.singleMulti_into_matrix(dynamic_question, matching_question)

    def answer_match(self, dynamic_question, questions):
        matching_question = next((question for question in questions \
                            if question.id == dynamic_question.carry_forward_answers_id), None)
        if matching_question.type == 'CompositeMatrix':
            dynamic_question.question_order = matching_question.question_order
            for question in dynamic_question.questions:
                question.response_order = matching_question.question_order
                for sub_question in matching_question.questions:
                    sub_question.add_dynamic_response(sub_question.prompt, sub_question.code)
        elif matching_question.type == 'CompositeMultipleSelect':
            dynamic_question.question_order = matching_question.question_order
            for question in dynamic_question.questions:
                question.response_order = matching_question.question_order
                for sub_question in matching_question.questions:
                    question.add_dynamic_response(sub_question.prompt, sub_question.code)
        elif matching_question.type == 'MC':
            for question in dynamic_question.questions:
                question.response_order = matching_question.response_order
                for response in matching_question.responses:
                    question.add_dynamic_response(response.label, response.value)

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

    def constantsum_match(self, dynamic_question, questions):
        matching_question = next((question for question in questions \
                             if question.id == dynamic_question.carry_forward_question_id), None)
        if matching_question.type == "CompositeMultipleSelect":
            dynamic_question.question_order = matching_question.question_order
            for current_q in matching_question.questions:
                sub_question = question.Question()
                sub_question.name ='%s_x%s' % (dynamic_question.name, current_q.code)
                sub_question.id = '%s_%s' % (dynamic_question.name, current_q.code)
                sub_question.code = current_q.code
                sub_question.prompt = current_q.prompt
                for response in current_q.responses:
                    sub_question.add_dynamic_response(response.label, response.value)
                dynamic_question.add_question(sub_question)

    def matrix_into_matrix(self, dynamic_matrix, matching_matrix):
        dynamic_matrix.question_order = matching_matrix.question_order
        for current_q in matching_matrix.questions:
            sub_question = question.Question()
            sub_question.name ='%s_x%s' % (dynamic_matrix.name, current_q.code)
            sub_question.id = '%s_%s' % (dynamic_matrix.name, current_q.code)
            sub_question.code = current_q.code
            sub_question.prompt = current_q.prompt
            for response in current_q.responses:
                sub_question.add_dynamic_response(response.label, response.value)
            dynamic_matrix.add_question(sub_question)

    def multiselect_into_matrix(self, dynamic_matrix, matching_multiselect):
        dynamic_matrix.question_order = matching_multiselect.question_order
        for current_q in matching_multiselect.questions:
            sub_question = question.Question()
            sub_question.name = '%s_x%s' % (dynamic_matrix.name, current_q.code)
            sub_question.id = '%s_%s' % (dynamic_matrix.id, current_q.code)
            sub_question.code = current_q.code
            sub_question.prompt = current_q.prompt
            for sub_responses in dynamic_matrix.temp_responses:
                sub_question.add_dynamic_response(sub_responses.label, sub_responses.value)
            dynamic_matrix.add_question(sub_question)

    def singleMulti_into_matrix(self, dynamic_matrix, matching_MC):
        dynamic_matrix.question_order = matching_MC.response_order
        for response in matching_MC.responses:
            sub_question = question.Question()
            sub_question.name = '%s_x%s' % (dynamic_matrix.name, response.value)
            sub_question.id = '%s_%s' % (dynamic_matrix.id, response.value)
            sub_question.code = response.value
            sub_question.prompt = response.label
            for sub_response in dynamic_matrix.temp_responses:
                sub_question.add_dynamic_response(sub_response.label, sub_response.value)
            dynamic_matrix.add_question(sub_question)

    def matrix_into_multiselect(self, multiselect, matrix):
        multiselect.question_order = matrix.question_order
        for current_q in matrix.questions:
            sub_question = question.Question()
            sub_question.name = '%s_x%s' % (matrix.name, current_q.code)
            sub_question.id = '%s_%s' % (multiselect.id, current_q.code)
            sub_question.code = current_q.code
            sub_question.prompt = current_q.prompt
            sub_question.add_dynamic_response('1', 1)
            multiselect.add_question(sub_question)

    def multiselect_into_multiselect(self, dynamic_multi, matching_multi):
        dynamic_multi.question_order = matching_multi.question_order
        for sub_question in matching_multi.questions:
            current_q = question.Question()
            current_q.name = '%s_x%s' % (dynamic_multi.name, sub_question.code)
            current_q.id = '%s_%s' % (dynamic_multi.id, sub_question.code)
            current_q.code = sub_question.code
            current_q.prompt = sub_question.prompt
            for response in sub_question.responses:
                current_q.add_dynamic_response(response.label, response.value)
            dynamic_multi.add_question(current_q)

    def singleMulti_into_multiselect(self, multiselect_question, matching_MC):
        multiselect_question.question_order = matching_MC.response_order
        for response in matching_MC.responses:
            sub_question = question.Question()
            sub_question.name = '%s_x%s' % (multiselect_question.name, response.value)
            sub_question.id = '%s_%s' % (multiselect_question.id, response.value)
            sub_question.code = response.value
            sub_question.prompt = response.label
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
            dynamic_MC.add_dynamic_response(response.label, response.value)

class MLStripper(parser.HTMLParser):
    def __init__(self):
        self.reset()
        self.strict= False
        self.convert_charrefs = True
        self.fed = []
    
    def handle_data(self, data):
        self.fed.append(data)
    
    def get_data(self):
        return ' '.join(self.fed)

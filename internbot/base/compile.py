from qsf_parser import QSFSurveyParser, QSFBlockFlowParser, QSFBlocksParser, QSFQuestionsParser
import json

class QSFSurveyCompiler(object):

    def __init__(self):
        self.survey_parser = QSFSurveyParser()
        self.blockflow_parser = QSFBlockFlowParser()
        self.blocks_parser = QSFBlocksParser()
        self.questions_parser = QSFQuestionsParser()

    def compile(self, path_to_qsf):
        try:
            qsf_json = self.parse_json(path_to_qsf)
            survey = self.compile_survey(qsf_json)
            return survey
        except UnicodeEncodeError, e:
            character = str(e)
            character = character.split(" ")
            error_word = character[5]
            print 'Remove word: %s' % (error_word)

    def compile_xtabs(self, path_to_qsf, path_to_text):    
        try:
            qsf_json = self.parse_json(path_to_qsf)
            questions = self.compile_questions(qsf_json)
            file = open(path_to_text, "w+")
            file.write(self.xtabs(questions))
        except UnicodeEncodeError, e:
            character = str(e)
            character = character.split(" ")
            error_word = character[5]
            print 'Remove word: %s' % (error_word)
            
    def parse_json(self, path_to_qsf):
        with open(path_to_qsf) as file:
            qsf_file = json.load(file)
        return qsf_file
        
    def compile_survey(self, qsf_json):
        survey = self.survey_parser.parse(qsf_json['SurveyEntry'])
        blocks = self.compile_blocks(qsf_json)
        questions = self.compile_questions(qsf_json)
        for block in blocks:
            survey.add_block(block)
        survey.add_questions(questions)
        return survey

    def compile_blocks(self, qsf_json):
        block_ids = self.blockflow_parser.parse(self.find_element('FL', qsf_json))
        blocks = self.blocks_parser.parse(self.find_element('BL', qsf_json))
        blocks.sort(block_ids)
        return blocks
        
    def compile_questions(self, qsf_json):    
        return self.questions_parser.parse(self.find_elements('SQ', qsf_json))
    
    def find_element(self, element_name, qsf_json):
        return next(iter(self.find_elements(element_name, qsf_json)), None)

    def find_elements(self, element_name, qsf_json):
        elements = qsf_json['SurveyElements']
        return [element for element in elements if element['Element'] == element_name]
    
    def xtabs(self, questions):
        result = ''
        grouped_questions = []
        grouped_name = []
        for question in questions:
            if question.type == 'MC' and question.subtype in ['SAVR','SAHR','DL']:
                result += 'VARIABLE LEVEL  %s(ORDINAL).\n' % question.name
            elif question.type == 'Slider':
                result += 'VARIABLE LEVEL  %s_1(ORDINAL).\n' % question.name
            elif question.type == 'Composite' and question.subtype == 'SingleAnswer':
                grouped_questions.append(self.xtabs_composite(question, grouped_name))    
            elif question.type == 'MC' and question.subtype == 'MAVR':
                grouped_questions.append(self.xtabs_mc_multiple(question, grouped_name))
        result += self.xtabs_group(grouped_questions, grouped_name, result)
        return result           
                
    def xtabs_composite(self, question, name):
        label = '$%s' % question.name
        name.append(label)
        result = "  /MDGROUP NAME=$%s LABEL='Select all that apply.'" % question.name
        result += "CATEGORYLABELS=VARLABELS\n"
        result+= "    VARIABLES="
        for sub_question in question.questions:
            result += "%s " % sub_question.name
        result += "VALUE=1\n"
        return result
        
    def xtabs_mc_multiple(self, question, name):
        label = '$%s' % question.name
        name.append(label)
        result = "  /MDGROUP NAME=$%s LABEL='Select all that apply.'" % question.name
        result += "CATEGORYLABELS=VARLABELS\n"
        result+= "    VARIABLES="
        for response in question.responses:
            result += "%s_x%s " % (question.name, response.code)
        result += "VALUE=1\n"
        return result
        
    def xtabs_group(self, grouped_questions, grouped_name, result):
        result += 'EXECUTE.\n\nMRSETS\n'
        for item in grouped_questions:
            result += item
        result += '  /DISPLAY NAME=['
        iteration = 1
        for name in grouped_name:    
            result += name
            if iteration < len(grouped_name):
                result += " "
            iteration += 1    
        result += '].'
        return result

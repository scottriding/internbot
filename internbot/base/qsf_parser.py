import json
from survey import Survey
from block import Blocks, Block

class QSFParser(object):

    def __init__(self):
        self.survey_parser = QSFSurveyParser()
        self.blocks_parser = QSFBlocksParser()
#         self.question_parser = QSFQuestionParser()


    def parse(self, path_to_qsf):
        qsf_json = self.parse_json(path_to_qsf)
        survey = self.survey_parser.parse(qsf_json['SurveyEntry'])
        blocks = self.blocks_parser.parse(self.find_element('BL', qsf_json))
        for block in blocks:
            survey.add_block(block)
#         questions = question_parser.parse(qsf_json[''])
#         for question in questions:
#             survey.add_question(question)
        return survey



#     def create_survey(self):
#         file = self.input_qsf()
#         name = file[u'SurveyEntry'][u'SurveyName']
#         elements = file[u'SurveyElements']
#         for element in elements:
#             if 'BL' in element[u'Element']:
#                 block_details = element
#             elif 'SQ' in element[u'Element']:
#                 question_details = element
#             elif 'QC' in element[u'Element']:
#                 question_count = element[u'SecondaryAttribute']
#         survey = Survey(name, block_details, question_details, question_count)
#         self.output_survey(survey)

    def parse_json(self, path_to_qsf):
        with open(path_to_qsf) as file:
            qsf_file = json.load(file)
        return qsf_file

    def find_element(self, element_name, qsf_json):
        elements = qsf_json['SurveyElements']
        return next((element for element in elements if element['Element'] == element_name), None)

class QSFSurveyParser(object):

    def parse(self, survey_element):
        return Survey(survey_element['SurveyName'])

class QSFBlocksParser(object):

    def parse(self, blocks_element):
        blocks = Blocks()
        for block_element in blocks_element['Payload']:
            block = Block(block_element['Description'])
            blocks.add(block)
        return blocks


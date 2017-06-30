import json
from survey import Survey
    
class Parser (object):
    def __init__ (self, path_to_qsf, path_to_output):
        self.path_to_qsf = path_to_qsf
        self.path_to_output = path_to_output
    
    def create_survey(self):
        file = self.input_qsf()
        description = file[u'SurveyEntry']
        elements = file[u'SurveyElements']
        block_details = []
        question_details = []
        for element in elements:
            if 'BL' in element[u'Element']:
                block_details = element
            elif 'SQ' in element[u'Element']:
                question_details = element
            elif 'QC' in element[u'Element']:
                question_count = element
        survey = Survey(description[u'SurveyName'], block_details, question_details, question_count)
        self.output_survey(survey)
        
    def input_qsf(self):
        with open(self.path_to_qsf) as file:
            survey_file = json.load(file)
        return survey_file
        
    def output_survey(self, survey):
        print survey
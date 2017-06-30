import json
from survey import Survey
    
class Parser (object):
    def __init__ (self, path_to_qsf, path_to_output):
        self.path_to_qsf = path_to_qsf
        self.path_to_output = path_to_output
        
    def input_qsf(self):
        with open(self.path_to_qsf) as file:
            survey_file = json.load(file)
        self.create_survey(survey_file)    
    
    def create_survey(self, file):
        description = file[u'SurveyEntry']
        elements = file[u'SurveyElements']
        survey = Survey(description[])
        
    def output_survey(self):
        pass
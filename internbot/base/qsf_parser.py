import json
    
class Parse_qsf(object):
    def __init__ (self, path_to_qsf, path_to_output):
        self.path_to_qsf = path_to_qsf
        self.path_to_output = path_to_output
        
    def input_survey(self, input_path):
        with open(input_path) as file:
            survey_file = json.load(file)
        # insert create a survey object    
            
    def output_survey(self, output_path):
        pass
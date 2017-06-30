from block import Blocks

class Survey(object):
    def __init__(self, survey_name, survey_block_details, survey_question_details, survey_question_count):
        self.name = survey_name
        self.block_details = survey_block_details
        self.question_details = survey_question_details
        self.question_count = survey_question_count
        
    def add_block(self):
        pass
        
    def __repr__ (self):
        result = ""
        result += "%s: %s" % (self.name, str(self.question_count))
        return result
        
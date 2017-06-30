from block import Blocks

class Survey(object):
    def __init__(self, survey_name, survey_q_count, survey_bl_count):
        self.name = survey_name
        self.q_count = survey_q_count
        self.bl_count = survey_bl_count

    def add_block(self):
        pass
        
    def __repr__ (self):
        result = ""
        result += "%s: %s/%s" % (self.name, str(self.blo_count), str(self.q_count))
        return result
        
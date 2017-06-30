from block import Blocks

class Survey(object):

    def __init__(self, survey_name):
        self.blocks = Blocks()
        self.name = survey_name

    def add_block(self, block):
        self.blocks.add(block)

    def __repr__ (self):
        result = ''
        result += "Survey: %s\n" % (self.name)
        result += str(self.blocks)
        return result

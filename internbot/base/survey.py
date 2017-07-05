from block import Blocks

class Survey(object):

    def __init__(self, survey_name):
        self.__blocks = Blocks()
        self.name = survey_name

    def add_block(self, block):
        self.__blocks.add(block)

    def add_questions(self, questions):
        for question in questions:
            block = self.__blocks.find_by_assigned_id(question.id)
            if block is not None:
                block.add_question(question)

    def __repr__ (self):
        result = ''
        result += "Survey: %s\n" % (self.name)
        result += str(self.__blocks)
        return result

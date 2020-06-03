from model.base import block

class Survey(object):

    def __init__(self, survey_name):
        self.__blocks = block.Blocks()
        self.name = survey_name
        self.__questions = []
        self.__scoring = []
        
    @property
    def blocks(self):
        return self.__blocks

    def add_block(self, block):
        self.__blocks.add(block)

    def add_questions(self, questions):
        for question in questions:
            block = self.__blocks.find_by_assigned_id(question.id)
            if block is not None:
                block.add_question(question)

    def add_scores(self, scores):
        self.__scoring.add(scores)

    @property
    def blocks(self):
        return self.__blocks

    def __repr__ (self):
        result = ''
        result += "Survey: %s\n" % (self.name)
        result += str(self.__blocks)
        return result
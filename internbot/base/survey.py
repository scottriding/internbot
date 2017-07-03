from block import Blocks

class Survey(object):

    def __init__(self, survey_name):
        self.__blocks = Blocks()
        self.name = survey_name

    def add_block(self, block):
        self.__blocks.add(block)

    def update_questions(self, questions):
        for question in questions:
            if self.__blocks.question_id_exists(question.id):
                block = self.__blocks.find_by_question_id(question.id)
                block.update_question(question)

    def __repr__ (self):
        result = ''
        result += "Survey: %s\n" % (self.name)
        result += str(self.__blocks)
        return result

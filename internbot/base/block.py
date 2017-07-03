from question import Questions

class Blocks(object):

    def __init__(self):
        self.__blocks = []

    def add(self, block):
        self.__blocks.append(block)

    def question_id_exists(self, question_id):
        return self.find_by_question_id(question_id) is not None

    def find_by_question_id(self, question_id):
        for block in self.__blocks:
            if block.find_by_id(question_id) is not None:
                return block
        return None

    def update_question(self, question):
        block = self.find_by_question_id(question.id)
        block.update_question(question)

    def __repr__(self):
        result = ''
        for block in self.__blocks:
            result += "\t%s\n" % str(block)
        return result

    def __iter__(self):
        return(iter(self.__blocks))

class Block(object):

    def __init__ (self, block_name):
        self.__questions = Questions()
        self.name = block_name

    def add_question(self, question):
        self.__questions.add(question)

    def update_question(self, question):
        self.__questions.replace(question)

    def find_by_id(self, question_id):
        return next((question for question in self.__questions if question.id == question_id), None)

    def __repr__(self):
        result = "Block: %s\n" % (self.name)
        result += str(self.__questions)
        return result

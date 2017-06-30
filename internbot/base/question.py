# add question type

from response import Responses

class Questions(object):

    def __init__(self):
        self.__questions = []

    def add(self, question):
        self.__questions.append(question)

    def get_by_name(self, question_name):
        pass

    def get_by_id(self, question_id):
        pass

    def __len__(self):
        return len(self.__questions)

    def __iter__(self):
        return iter(self.__questions)

    def __repr__(self):
        result = ''
        for question in self.__questions:
            result += "\t\t%s\n" % str(question)
        return result


class Question(object):

    def __init__(self, id=''):
        self.__responses = Responses()
        self.id = id

    @property
    def id(self):
        return self.__id

    @id.setter
    def id(self, id):
        self.__id = str(id)

    @property
    def name(self):
        return self.__name

    @name.setter
    def name(self, name):
        self.__name = str(name)

    @property
    def prompt(self):
        return self.__prompt

    @prompt.setter
    def prompt(self, prompt):
        self.__prompt = str(prompt)

    @property
    def n(self):
        return self.__n

    @n.setter
    def n(self, n):
        self.__n = int(n)

    @property
    def responses(self):
        return self.__responses

    def add_response(self, response, freq):
        self.__responses.add(response, freq)

    def __repr__(self):
        return "Question: %s" % self.id


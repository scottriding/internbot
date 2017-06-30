# add question type

from response import Responses

class Questions(object):

    def __init__(self):
        self.__questions = []

    def add(self, question):
        self.__questions.append(question)

    def replace(self, question):
        location = self.location_by_id(question.id)
        self.__questions[location] = question

    def location_by_id(self, question_id):
        next((i for i, question in enumerate(self.__questions) if question.id == question_id), None)

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

    def add_response(self, response, code=None):
        self.__responses.add(response, code)

    def __repr__(self):
        result = ''
        result += "Question: %s" % self.id
        result += str(self.__responses)
        return result


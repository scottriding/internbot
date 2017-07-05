# add question type

from response import Responses
import re

class Questions(object):

    def __init__(self):
        self.__questions = []

    def add(self, question):
        self.__questions.append(question)

    def sort(self, question_id_order):
        sorter = QuestionSorter(question_id_order)
        self.__questions = sorter.sort(self.__questions)

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
    def type(self):
        return self.__type

    @type.setter
    def type(self, type):
        self.__type = str(type)

    @property
    def subtype(self):
        return self.__subtype

    @subtype.setter
    def subtype(self, subtype):
        self.__subtype = str(subtype)

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
        result += "Question: %s\n" % self.id
        result += str(self.__responses)
        return result


class QuestionSorter(object):

    def __init__(self, question_order):
        self.__order = question_order

    def sort(self, questions):
        return sorted(questions, cmp=self.compare)

    def compare(self, question1, question2):
        id1_components = re.match('(QID\d+)(_\d+)?', question1.id)
        id2_components = re.match('(QID\d+)(_\d+)?', question2.id)
        id1_location = self.__order.index(id1_components.group(1))
        id2_location = self.__order.index(id2_components.group(1))
        if id1_location > id2_location:
            return 1
        elif id1_location < id2_location:
            return -1
        elif id1_components.group(2) > id2_components.group(2):
            return 1
        elif id1_components.group(2) < id2_components.group(2):
            return -1
        else:
            return 0


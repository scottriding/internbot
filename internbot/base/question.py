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

    def __init__(self):
        self.__responses = Responses()
        self.__response_order = []
        self.has_carry_forward_responses = False
        self.has_carry_forward_prompts = False
        self.subtype = ''

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
    def has_carry_forward_responses(self):
        return self.__has_carry_forward_responses

    @has_carry_forward_responses.setter
    def has_carry_forward_responses(self, has_carry_forward_responses):
        self.__has_carry_forward_responses = bool(has_carry_forward_responses)

    @property
    def has_carry_forward_prompts(self):
        return self.__has_carry_forward_prompts

    @has_carry_forward_prompts.setter
    def has_carry_forward_prompts(self, has_carry_forward_prompts):
        self.__has_carry_forward_prompts = bool(has_carry_forward_prompts)

    @property
    def carry_forward_question_id(self):
        return self.__carry_forward_question_id

    @carry_forward_question_id.setter
    def carry_forward_question_id(self, carry_forward_question_id):
        self.__carry_forward_question_id = str(carry_forward_question_id)

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

    @property
    def response_order(self):
        return self.__response_order

    @response_order.setter
    def response_order(self, response_order):
        for response_location in response_order:
            self.__response_order.append(str(response_location))

    def spss_name(self):
        if self.type == 'MC' and self.subtype in ['SAVR','SAHR']:
            return self.name
        elif self.type == 'Matrix' and self.subtype == 'SingleAnswer':
            return self.name
        elif self.type == 'MC' and self.subtype == 'MAVR':
            names = []
            for response in self.responses:
                names.append('%s_%s' % (self.name, response.code))
            return names
        elif self.type == 'DB':
            return self.name
        else:
            return 'SPSS name not known'

    def add_response(self, response, code=None):
        self.__responses.add(response, code)
        self.__responses.sort(self.response_order)

    def __repr__(self):
        result = ''
        result += "Question: %s\n" % self.spss_name()
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


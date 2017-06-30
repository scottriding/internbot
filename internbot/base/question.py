from response import Responses
from collections import OrderedDict

class Questions(object):

    def __init__(self, questions_data = []):
        self.__questions = OrderedDict()
        for question_data in questions_data:
            self.add(question_data)

    def add(self, question_data):
        if self.already_exists(question_data['name']):
            question = self.get(question_data['name'])
            question.add_response(question_data['response'], question_data['frequency'])
        else:
            self.add_new(question_data)

    def add_new(self, question_data):
        q = Question(question_data['name'], question_data['prompt'], question_data['n'])
        q.add_response(question_data['response'], question_data['frequency'])
        self.__questions[q.name] = q

    def get(self, question_name):
        return self.__questions.get(question_name)

    def already_exists(self, question_name):
        if self.get(question_name) is None:
            return False
        else:
            return True

    def __len__(self):
        return len(self.__questions)

    def __repr__(self):
        result = ''
        for name, question in self.__questions.iteritems():
            result += str(question)
            result += '\n'
        return result

    def __iter__(self):
        return iter(self.__questions.values())


class Question(object):

    def __init__(self, name, prompt, n):
        self.name      = name
        self.prompt    = prompt
        self.n         = n
        self.__responses = Responses()

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
        result = ""
        result += "%s: %s" % (self.name, self.prompt)
        for response, freq in self.responses.iteritems():
            result += ' %s: %s' % (response, str(freq))
        return result

from collections import OrderedDict

class Questions(object):

    def __init__(self, questions_data = []):
        self.__questions = OrderedDict()
        for question_data in questions_data:
            self.add(question_data)

    def add(self, question_data):
        display_logic = question_data['logic']
        if self.already_exists(question_data['variable']):
            question = self.get(question_data['variable'])
            question.add_response(question_data['label'], question_data['percent'], question_data['n'])
            if display_logic != "":
                question.add_display(display_logic)
        else:
            question = Question(question_data['variable'], question_data['prompt'])
            question.add_response(question_data['label'], question_data['percent'], question_data['n'])
            if display_logic != "":
                question.add_display(display_logic)
            self.__questions[question.name] = question

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

    def list_names(self):
        result = self.__questions.keys()
        return result

    def prompt(self, question_name):
        result = self.__questions[question_name].prompt
        return result

    def n(self, question_name):
        result = self.__questions[question_name].n
        return result

    def return_responses(self, question_name):
        result = self.__questions[question_name].return_response
        return result

    def display_logic(self, question_name):
        result = self.__questions[question_name].display_logic
        return result


class Question(object):

    def __init__(self, name, prompt):
        self.name = name
        self.prompt = prompt
        self.n = 0
        self.responses = OrderedDict()
        self.__display_logic = ""

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
        self.__n = int(float(n))

    def add_response(self, response, freq, n):
        self.responses[str(response)] = float(freq)
        self.n += float(n)

    def add_display(self, logic):
        self.__display_logic = logic

    @property
    def return_response(self):
        return self.responses

    @property
    def display_logic(self):
        return self.__display_logic

    def __repr__(self):
        result = ""
        result += "%s: %s" % (self.name, self.prompt)
        for response, freq in self.responses.iteritems():
            result += ' %s: %s' % (response, str(freq))
        return result

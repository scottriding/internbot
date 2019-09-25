from collections import OrderedDict

class CSVQuestions(object):

    def __init__(self, questions_data=[]):
        self.__questions = OrderedDict()
        for question_data in questions_data:
            self.add(question_data)

    def add(self, question_data, groups):
        question_name = question_data["variable"]

        question_prompt = question_data['prompt']
        question_response = question_data['label']
        question_pop = question_data['n']
        question_stat = question_data['stat']
        if self.already_exists(question_name):
            question = self.get(question_name)
            if question_response != "":
                question.add_response(question_response, question_data, question_pop, groups)
            if question_stat != "":
                question.add_stat(question_stat)
        else:
            question = CSVQuestion(question_name, question_prompt)
            if question_response != "":
                question.add_response(question_response, question_data, question_pop, groups)
            if question_stat != "":
                question.add_stat(question_stat)
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

    def list_names(self):
        result = self.__questions.keys()
        return result

    def __repr__(self):
        result = ''
        for name, question in self.__questions.iteritems():
            result += str(question)
            result += '\n'
        return result


class CSVQuestion(object):

    def __init__(self, name, prompt):
        self.__name = name
        self.__prompt = prompt
        self.__n = 0
        self.__responses = []
        self.__stat = ""

    @property
    def name(self):
        return self.__name

    @property
    def prompt(self):
        return self.__prompt

    @property
    def n(self):
        return self.__n

    @property
    def responses(self):
        return self.__responses

    @property
    def stat (self):
        return self.__stat

    def add_response(self, response_name, response_data, response_pop, groups):
        self.__n += int(response_pop)
        self.__responses.append(CSVResponse(response_name, response_data, groups))

    def add_stat(self, stat):
        self.__stat = stat


class CSVResponse(object):

    def __init__(self, label, frequency_data, groups):
        self.__name = label
        self.__frequencies = OrderedDict() 
        if len(groups) > 0:
            for group in groups:
                round_col = "result %s" % group
                if frequency_data[round_col] != "":
                    self.__frequencies[group] = frequency_data[round_col]
        else:
            round_col = "result"
            self.__frequencies[0] = frequency_data[round_col] 

    @property
    def name(self):
        return self.__name

    @property
    def frequencies(self):
        return self.__frequencies

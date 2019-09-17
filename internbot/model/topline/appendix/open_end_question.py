class OpenEndQuestion(object):

    def __init__(self, name, prompt):
        self.__name = name
        self.__prompt = prompt
        self.__responses = []
        self.__response_count = 0

    def add_response(self, response):
        if response != "":
            self.__responses.append(response)
            self.__response_count += 1
            self.__responses.sort()

    @property
    def name(self):
        return self.__name

    @property
    def prompt(self):
        return self.__prompt

    @property
    def responses(self):
        return self.__responses

    @property
    def response_count(self):
        return self.__response_count

from operator import attrgetter
from sorter import ResponseSorter

class Responses(object):

    def __init__(self):
        self.__responses = []

    def add(self, response, code=None):
        self.__responses.append(
            Response(response, code)
        )

    def add_text(self, response):
        self.__responses.append(
            TextResponse(response)
        )

    def sort(self, response_order):
        sorter = ResponseSorter(response_order)
        self.__responses = sorter.sort(self.__responses)

    def __len__(self):
        return len(self.__responses)

    def __iter__(self):
        return(iter(self.__responses))

    def __repr__(self):
        result = ''
        for response in self.__responses:
            result += "\t\t\t%s\n" % str(response)
        return result

class Response(object):

    def __init__ (self, response, code=None):
        self.response = response
        self.code = code
        self.__has_frequency = False

    @property
    def type(self):
        return 'Response'

    @property
    def response(self):
        return self.__response

    @response.setter
    def response(self, response):
        self.__response = str(response)

    @property
    def code(self):
        return self.__code

    @code.setter
    def code(self, code):
        self.__code = str(code)

    @property
    def frequency(self):
        return self.__frequency

    @frequency.setter
    def frequency(self, frequency):
        self.__frequency = float(frequency)
        self.__has_frequency = True

    @property
    def has_frequency(self):
        return self.__has_frequency

    def __repr__(self):
        result = ""
        result += "%s: %s" % (self.code, self.response)
        return result

class TextResponse(Response):

    def __init__(self, response):
        super(TextResponse, self).__init__(response)

    @property
    def type(self):
        return 'TextResponse'

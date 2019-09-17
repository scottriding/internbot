from model.base import sorter

from operator import attrgetter
from collections import OrderedDict

class Responses(object):

    def __init__(self):
        self.__responses = []

    def add(self, response, code=None):
        self.__responses.append(
            Response(response, code)
        )

    def add_dynamic(self, response, code=None):
        response = Response(response, code)
        response.is_dynamic = True
        self.__responses.append(response)

    def add_text(self, response):
        self.__responses.append(
            TextResponse(response)
        )

    def add_NA(self):
        self.__responses.append(
            NAResponse()
        )

    def get_NA(self):
        return next((response for response in self.__responses \
                     if response.type == 'NAResponse'), None)

    def get_first(self):
        return self.__responses[0]

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
        self.__is_dynamic = False
        self.__frequencies = OrderedDict()

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
    def frequencies(self):
        return self.__frequencies

    @frequencies.setter
    def frequencies(self, frequency):
        self.__frequencies = frequency
        self.__has_frequency = True

    @property
    def has_frequency(self):
        return self.__has_frequency

    @property
    def is_dynamic(self):
        return self.__is_dynamic

    @is_dynamic.setter
    def is_dynamic(self, type):
        self.__is_dynamic = bool(type)

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

class NAResponse(Response):

    def __init__(self):
        super(NAResponse, self).__init__('NA')

    @property
    def type(self):
        return 'NAResponse'

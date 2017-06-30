# add appearance order

from operator import attrgetter

class Responses(object):

    def __init__(self):
        self.__responses = []

    def add(self, response, frequency):
        self.__responses.append(
            Response(response, frequency)
        )
        self.sort()

    def sort(self):
        self.__responses.sort(
            key = attrgetter('frequency'),
            reverse = True
        )

    def __iter__(self):
        return(iter(self.__responses))


class Response(object):

    def __init__ (self, response, frequency):
        self.response = response
        self.frequency = frequency

    @property
    def response(self):
        return self.__response

    @response.setter
    def response(self, response):
        self.__response = str(response)

    @property
    def frequency(self):
        return self.__frequency

    @frequency.setter
    def frequency(self, frequency):
        self.__frequency = float(frequency)

    def __repr__(self):
        result = ""
        result += "%s: %s" % (self.response, str(self.frequency))
        return result

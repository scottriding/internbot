from model.base import sorter
from model.base import frequency

from operator import attrgetter
from collections import OrderedDict

class Responses(object):

    def __init__(self):
        self.__responses = []

    def add(self, label, value=None):
        new = Response(label, value)
        self.__responses.append(
            new
        )
        return new

    def add_dynamic(self, label, value=None):
        response = Response(label, value)
        response.is_dynamic = True
        self.__responses.append(response)

    def sort(self, response_order):
        response_sorter = sorter.ResponseSorter(response_order)
        self.__responses = response_sorter.sort(self.__responses)

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

    def __init__ (self, label, value=None):
        self.__label = label
        self.__value = value
        self.__is_dynamic = False
        self.__frequencies = frequency.Frequencies()

    @property
    def type(self):
        return 'Response'

    @property
    def label(self):
        return self.__label

    @label.setter
    def label(self, label):
        self.__label = str(label)

    @property
    def value(self):
        return self.__value

    @value.setter
    def value(self, value):
        self.__value = str(value)

    @property
    def frequencies(self):
        return self.__frequencies

    @frequencies.setter
    def frequencies(self, frequencies):
        self.__frequencies = frequencies

    @property
    def is_dynamic(self):
        return self.__is_dynamic

    @is_dynamic.setter
    def is_dynamic(self, type):
        self.__is_dynamic = bool(type)

    def add_frequency(self, result, population, stat, group="Basic"):
        self.__frequencies.add(result, population, stat, group)

    def __repr__(self):
        result = ""
        result += "%s: %s" % (self.__value, self.__label)
        return result

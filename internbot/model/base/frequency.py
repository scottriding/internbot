from collections import OrderedDict

class Frequencies(object):

    def __init__(self):
        self.__frequencies = OrderedDict()
        self.__groups = []

    @property
    def frequencies(self):
        return self.__frequencies

    @property
    def groups(self):
        return self.__groups

    def add(self, result, population, stat, group="basic"):
        if self.__frequencies.get(group) is not None:
            group_freqs = self.__frequencies[group]
            groups_freqs.append(Frequency(result, population, stat))
        else:
            self.__frequencies[group] = [Frequency(result, population, stat)]
            self.__groups.append(group)

class Frequency(object):

    def __init__(self, result, population, stat):
        self.__result = result
        self.__population = population
        self.__stat = stat

    @property
    def result(self):
        return self.__result

    @property
    def population(self):
        return self.__population

    @property
    def stat(self):
        return self.__stat

    def __repr__(self):
        result = ''
        result += "%s: %s" % (self.result, self.population)
        return result

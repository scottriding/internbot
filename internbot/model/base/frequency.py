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
			self.__frequencies[group] = Frequency(result, population, stat)
			self.__groups.append(group)

	def __len__(self):
		return len(self.__frequencies)

	def __iter__(self):
		return(iter(self.__frequencies))

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

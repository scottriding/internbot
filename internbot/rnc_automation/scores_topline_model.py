from collections import OrderedDict

class ScoreToplineModels(object):

    def __init__(self, round_count):
        self.rounds = round_count
        self.__models = OrderedDict()

    def add_model(self, model_data):
        model_name = model_data["Model"]
        model_desc = model_data["Survey question reference"]
        
        if self.already_exists(model_name):
            model = self.__models.get(model_name)
            model.add_variable(model_data)
        else:
            new_model = ScoreToplineModel(self.rounds, model_name, model_desc)
            new_model.add_variable(model_data)
            self.__models[model_name] = new_model

    def already_exists(self, model_name):
        if self.__models.get(model_name) is None:
            return False
        else:
            return True

    def list_model_names(self):
        return self.__models.keys()

    def get_model(self, model_name):
        return self.__models.get(model_name)

class ScoreToplineModel(object):

    def __init__(self, rounds, model_name, model_desc):
        self.rounds = rounds
        self.__name = str(model_name)
        self.__description = str(model_desc)
        self.__variables = OrderedDict()

    def add_variable(self, model_data):
        model_variable = model_data["Variable"]

        weighted_frequencies = []
        unweighted_frequencies = []

        round_iteration = 1
        while round_iteration <= self.rounds:
            round_header_unweighted = "Round %s Frequency" % round_iteration
            round_header_weighted = "Round %s TOW Frequency" % round_iteration
            unweighted = model_data[round_header_unweighted]
            weighted = model_data[round_header_weighted]
            unweighted_frequencies.append(float(unweighted))
            weighted_frequencies.append(float(weighted))
            round_iteration += 1

        new_variable = Variable(weighted_frequencies, unweighted_frequencies)
        self.__variables[model_variable] = new_variable

    def list_variable_names(self):
        return self.__variables.keys()

    def get_variable(self, variable_name):
        return self.__variables.get(variable_name)

    @property
    def name(self):
        return self.__name

    @property
    def description(self):
        return self.__description

class Variable(object):

    def __init__ (self, weighted_frequencies, unweighted_frequencies):
        self.__weighted_frequencies = weighted_frequencies
        self.__unweighted_frequencies = unweighted_frequencies

    def round_weighted_freq(self, round):
        return self.__weighted_frequencies[round]

    def round_unweighted_freq(self, round):
        return self.__unweighted_frequencies[round]

    def weighted_frequencies(self):
        return self.__weighted_frequencies

    def unweighted_frequencies(self):
        return self.__unweighted_frequencies
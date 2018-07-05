from collections import OrderedDict

class IssueTrendedNetModel(object):

    def __init__(self, round_count):
        self.rounds = round_count
        self.__models = OrderedDict()

    def add_model(self, model_data):
        model_name = model_data["Model"]

        if self.already_exists(model_name):
            net_model = self.__models.get(model_name)
            net_model.add_field(model_data)
        else:
            new_model = IssueTrendedModel(self.rounds, model_name)
            new_model.add_field(model_data)
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

class IssueTrendedModel(object):

    def __init__(self, rounds, model_name):
        self.rounds = rounds
        self.__name = str(model_name)
        self.__fields = OrderedDict()

    def add_field(self, model_data):
        field_name = model_data["Field Name"]
        if self.already_exists(field_name):
            field = self.__fields.get(field_name)
            field.add_grouping(model_data)
        else:
            new_field = IssueTrendedField(self.rounds, field_name)
            new_field.add_grouping(model_data)
            self.__fields[field_name] = new_field

    def already_exists(self, field_name):
        if self.__fields.get(field_name) is None:
            return False
        else:
            return True

    def list_field_names(self):
        return self.__fields.keys()

    def get_field(self, field_name):
        return self.__fields.get(field_name)

    @property
    def name(self):
        return self.__name

class IssueTrendedField(object):

    def __init__(self, rounds, field_name):
        self.rounds = rounds
        self.__name = str(field_name)
        self.__groupings = OrderedDict()

    def add_grouping(self, field_data):
        grouping_name = field_data["Grouping"]
        count = field_data["Count"]

        frequencies = []

        round_iteration = 1
        while round_iteration <= self.rounds:
            frequency_header = "Round %s Frequency" % round_iteration
            frequency = float(field_data[frequency_header])
            frequencies.append(frequency)
            round_iteration += 1

        new_grouping = Grouping(count, frequencies)
        self.__groupings[grouping_name] = new_grouping

    def list_grouping_names(self):
        return self.__groupings.keys()

    def get_grouping(self, grouping_name):
        return self.__groupings.get(grouping_name)

    @property
    def name(self):
        return self.__name

class Grouping(object):

    def __init__(self, count, frequencies):
        self.__count = long(count)
        self.__frequencies = frequencies

    def round_frequency(self, round):
        return self.__frequencies[round]

    @property
    def frequencies(self):
        return self.__frequencies

    @property
    def count(self):
        return self.__count
from collections import OrderedDict

class ModelFileGenerator(object):

    def __init__(self, round_count):
        self.__models = OrderedDict()

    def add_model(self, model_data):
        net_model_name = model_data["Model"]
        model_name = model_data["Variable"]
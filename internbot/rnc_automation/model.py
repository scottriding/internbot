from collections import OrderedDict

class Models(object):

    def __init__(self, models_data = []):
        self.__models = OrderedDict()
        for model_data in models_data:
            self.add(model_data)

    def add(self, model_data):
        if self.already_exists(model_data['Model']):
            model = self.get(model_data['Model'])
            model.add_variable(model_data['Variable'], model_data['Frequency'])
        else:
            self.add_new(model_data)

    def add_new(self, model_data):
        new_model = Model(model_data['Model'], model_data['Description'])
        new_model.add_variable(model_data['Variable'], model_data['Frequency'])
        self.__models[new_model.name] = new_model

    def get(self, model_name):
        return self.__models.get(model_name)

    def already_exists(self, model_name):
        if self.get(model_name) is None:
            return False
        else:
            return True

    def __len__(self):
        return len(self.__models)

    def __repr__(self):
        result = ''
        for name, model in self.__models.iteritems():
            result += str(model)
            result += '\n'
        return result

    def list_names(self):
        result = self.__models.keys()
        return result

    def description(self, model_name):
        result = self.__models[model_name].description
        return result

    def return_variables(self, model_name):
        result = self.__models[model_name].return_variable
        return result


class Model(object):

    def __init__(self, name, description=None):
        self.name = name
        self.description = description
        self.variables = OrderedDict()

    @property
    def name(self):
        return self.__name

    @name.setter
    def name(self, name):
        self.__name = str(name)

    @property
    def description(self):
        return self.__description

    @description.setter
    def description(self, description):
        self.__description = str(description)

    def add_variable(self, variable, freq):
        self.variables[str(variable)] = float(freq)

    @property
    def return_variable(self):
        return self.variables

    def __repr__(self):
        result = ""
        result += "%s: %s" % (self.name, self.description)
        for variable, freq in self.variables.iteritems():
            result += ' %s: %s' % (variable, str(freq))
        return result

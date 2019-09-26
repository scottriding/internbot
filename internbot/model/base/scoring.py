class Scorings(object):
    
    def __init__(self):
        self.__scorings = []

    def add(self, scoring_id, scoring_name, scoring_desc=None):
        self.__scorings.append(Scoring(scoring_id, scoring_name, scoring_desc))

    def __len__(self):
        return len(self.__scorings)

    def __iter__(self):
        return(iter(self.__scorings))

    def __repr__(self):
        result = ''
        for scoring in self.__scorings:
            result += str(scoring)
        return result

class Scoring(object):

    def __init__ (self, id, name, description=None):
        self.__id = id
        self.__name = name
        self.__desc = description

    @property
    def id(self):
        return self.__id

    @property
    def name(self):
        return self.__name

    @property
    def description(self):
        return self.__desc

    def __repr__(self):
        result = ""
        result += "%s: %s" % (self.name, self.id)
        return result



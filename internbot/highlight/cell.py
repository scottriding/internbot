class Cells(object):

    def __init__(self):
        self.__cells = []

    def add(self, cell):
        self.__cells.append(cell)

class Cell(object):

    def __init__(self, row_name, column_name, location):
        self.row = row_name
        self.column = column_name
        self.location = location
        self.is_significant = False

    @property
    def row(self):
        return self.row

    @property
    def column(self):
        return self.column

    @property
    def is_significant(self):
        return self.is_significant
    @property
    def row_name(self):
        return self.row_name

    @row_name.setter
    def row_name(self, name):
        self.row_name = str(name)

    @property
    def column_name(self):
        return self.column_name

    @column_name.setter
    def column_name(self, name):
        self.column_name = str(name)

class FrequencyCell(Cell):

    def __init__(self, row, column):
        super(Cell, self).__init__(row, column)

    @property
    def type(self):
        return 'FrequencyCell'

    @property
    def frequency(self):
        return self.frequency

    @frequency.setter
    def frequency(self, frequency):
        self.frequency = str(frequency)

    @is_significant.setter
    def is_significant(self, boolean):
        self.is_significant = bool(boolean)

class PopulationCell(Cell):

    def __init__(self, row, column):
        super(Cell, self).__init__(row, column)

    @property
    def type(self):
        return 'PopulationCell'

    @property
    def population(self):
        return self.population

    @population.setter
    def population(self, population):
        self.population = str(population)

class SignificantMarker(Cell):

    def __init__(self, row, column):
        super(Cell, self).__init__(row, column)

    @property
    def type(self):
        return 'SignificantMarker'

    @is_significant.setter
    def is_significant(self, boolean):
        self.is_significant = bool(boolean)
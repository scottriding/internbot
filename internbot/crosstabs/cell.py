class Cells(object):

    def __init__(self):
        self.__cells = []

    def add(self, cell):
        self.__cells.append(cell)

    def cells_to_highlight(self):
        to_highlight = []
        significant_values = self.find_significant_values()
        for value in significant_values:
            match = self.match_cells(value)
            if match is not None:
                to_highlight.append(match)

    def find_significant_values(self):
        result = [cell for cell in self.__cells \
                                if cell.is_significant == True]
        return result

    def matching_cells(self, value_to_match):
        for cell in self.__cells:
            if cell.row == value_to_match.row and \
               cell.column == value_to_match.column and \
               cell.location != value_to_match.location:
                match = cell
        return match

    def __iter__(self):
        return iter(self.__cells)

class Cell(object):

    def __init__(self, row, column, location):
        self.__row = row
        self.__column = column
        self.__location = location
        self.__is_significant = False

    @property
    def row(self):
        return self.__row

    @property
    def column(self):
        return self.__column

    @property
    def location(self):
        return self.__location

    @property
    def is_significant(self):
        return self.__is_significant

    @is_significant.setter
    def is_significant(self, boolean):
        self.__is_significant = bool(boolean)

class FrequencyCell(Cell):

    def __init__(self, row, column, location):
        super(FrequencyCell, self).__init__(row, column, location)

    @property
    def type(self):
        return 'FrequencyCell'

    @property
    def frequency(self):
        return self.frequency

    @frequency.setter
    def frequency(self, frequency):
        self.frequency = str(frequency)

class PopulationCell(Cell):

    def __init__(self, row, column, location):
        super(PopulationCell, self).__init__(row, column, location)

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

    def __init__(self, row, column, location):
        super(SignificantMarker, self).__init__(row, column, location)

    @property
    def type(self):
        return 'SignificantMarker'

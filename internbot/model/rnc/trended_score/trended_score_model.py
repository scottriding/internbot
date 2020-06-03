from collections import OrderedDict

class TrendedModelWorkbooks(object):
    
    def __init__(self, rounds):
        self.__workbooks = OrderedDict()
        self.rounds = rounds
        self.__round_dates = []

    def add(self, workbook_details):
        ## grab row details
        workbook_name = workbook_details['Workbook']
        sheet_name = workbook_details['Sheet']
        field_name = workbook_details['FieldName']
        grouping_name = workbook_details['Grouping']
        count = int(workbook_details['Count'])

        frequencies = []

        round_iteration = 1
        while round_iteration <= self.rounds:
            round_header = "Round %s Frequency" % round_iteration
            date_header = "Round %s Date" % round_iteration
            frequency = workbook_details[round_header]
            date = workbook_details[date_header]
            frequencies.append(frequency)
            self.__round_dates.append(date)
            round_iteration += 1

        ## add or update
        if self.already_exists(workbook_name):
            workbook = self.__workbooks.get(workbook_name)
            workbook.add_sheet(sheet_name, field_name, grouping_name, count, frequencies)
        else:
            new_workbook = TrendedModelWorkbook(workbook_name)
            new_workbook.add_sheet(sheet_name, field_name, grouping_name, count, frequencies)
            self.__workbooks[workbook_name] = new_workbook

    def already_exists(self, workbook_name):
        if self.__workbooks.get(workbook_name) is None:
            return False
        else:
            return True

    def get_workbook(self, workbook_name):
        return self.__workbooks.get(workbook_name)

    def list_workbook_names(self):
        return self.__workbooks.keys()

    def round_date(self, round_number):
        return self.__round_dates[round_number - 1]

class TrendedModelWorkbook(object):
    
    def __init__(self, name):
        self.__name = name
        self.__sheets = OrderedDict()

    def add_sheet(self, sheet_name, field_name, grouping_name, count, frequencies):
        if self.already_exists(sheet_name):
            sheet = self.__sheets.get(sheet_name)
            sheet.add_field(field_name, grouping_name, count, frequencies)
        else:
            new_sheet = TrendedModelSheet(sheet_name)
            new_sheet.add_field(field_name, grouping_name, count, frequencies)
            self.__sheets[sheet_name] = new_sheet

    def already_exists(self, sheet_name):
        if self.__sheets.get(sheet_name) is None:
            return False
        else:
            return True

    def get_sheet(self, sheet_name):
        return self.__sheets.get(sheet_name)

    def list_sheet_names(self):
        return self.__sheets.keys()

    @property
    def name(self):
        return self.__name

class TrendedModelSheet(object):
    
    def __init__(self, name):
        self.__name = name
        self.__fields = OrderedDict()

    def add_field(self, field_name, grouping_name, count, frequencies):
        if self.already_exists(field_name):
            field = self.get_field(field_name)
            field.add_grouping(grouping_name, count, frequencies)
        else:
            new_field = TrendedModelField(field_name)
            new_field.add_grouping(grouping_name, count, frequencies)
            self.__fields[field_name] = new_field

    def already_exists(self, field_name):
        if self.__fields.get(field_name) is None:
            return False
        else:
            return True

    def get_field(self, field_name):
        return self.__fields.get(field_name)

    def list_field_names(self):
        return list(self.__fields.keys())

    @property
    def name(self):
        return self.__name

class TrendedModelField(object):
    
    def __init__(self, name):
        self.__name = name
        self.__groupings = OrderedDict()

    def add_grouping(self, grouping_name, grouping_count, grouping_frequencies):
        new_grouping = TrendedModelGrouping(grouping_name, grouping_count, grouping_frequencies)
        self.__groupings[grouping_name] = new_grouping

    def get_grouping(self, grouping_name):
        return self.__groupings.get(grouping_name)

    def list_grouping_names(self):
        return self.__groupings.keys()

    def grouping_count(self):
        return len(list(self.__groupings.keys()))

    @property
    def name(self):
        return self.__name

class TrendedModelGrouping(object):
    
    def __init__(self, name, count, frequencies):
        self.__name = name
        self.__count = float(count)
        self.__frequencies = frequencies

    @property
    def name(self):
        return self.__name

    @property
    def count(self):
        return self.__count

    @property
    def frequencies(self):
        return self.__frequencies

    def round_frequency(self, round):
        return self.__frequencies[round - 1]

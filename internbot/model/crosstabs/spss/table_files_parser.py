import csv

class TablesParser(object):

    def __init__(self):
        self.__tables = []

    def create_tables(self, path_to_tables):
        with open(path_to_tables, 'r') as table_file:
            file = csv.DictReader(table_file, quotechar = '"')
            for row in file:
                table = Table(row['TableIndex'], row['VariableName'], row['Title'], row['Base'])
                tables.append(table)

    @property
    def tables(self):
        return self.__tables

class Table(object):

    def __init__(self, index, name, prompt, base):
        self.__index = index
        self.__name = name
        self.__prompt = prompt
        self.__base = base

    @property
    def index(self):
        return self.__index

    @property
    def name(self):
        return self.__name

    @property
    def prompt(self):
        return self.__prompt

    @property
    def base(self):
        return self.__base
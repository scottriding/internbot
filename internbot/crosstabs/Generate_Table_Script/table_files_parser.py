import csv

class TablesParser(object):

    def __init__(self):
        self.questions = []
        self.titles = []
        self.bases = []

    def pull_table_names(self, path_to_tables):
        with open(path_to_tables, 'rb') as table_file:
            file = csv.DictReader(table_file, quotechar = '"')
            for row in file:
                self.questions.append(row['VariableName'])
        return self.questions

    def pull_table_titles(self, path_to_tables):
        with open(path_to_tables, 'rb') as table_file:
            file = csv.DictReader(table_file, quotechar = '"')
            for row in file:
                self.titles.append(row['Title'])
        return self.titles

    def pull_table_bases(self, path_to_tables):
        with open(path_to_tables, 'rb') as table_file:
            file = csv.DictReader(table_file, quotechar = '"')
            for row in file:
                self.bases.append(row['Base'])
        return self.bases
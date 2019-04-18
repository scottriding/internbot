import csv

class TableScriptGenerator(object):

    def __init__(self):
        self.grouped_question = []

    def compile_scripts(self, tables, banners, embedded_variables, filtering_variable, path_to_output):
        self.__filtering_variable = filtering_variable
        table_path = tables
        output_path = str(path_to_output) + '/table script.sps'
        output = open(output_path, "w+")
        script = self.write_script(table_path, path_to_output, banners, embedded_variables)
        output.write(script)
                
    def write_script(self, path_to_tables, path_to_output, banners, embedded_variables):
        result = ''
        with open(path_to_tables, 'rb') as table_file:
            column_specs = banners
            result += self.add_column_recode(column_specs, embedded_variables)
            file = csv.DictReader(table_file, quotechar = '"')
            for table in file:
                result += self.write_table(table, path_to_output)
                result += self.write_ctable(table, column_specs)
        return result

    def write_table(self, question, path_to_output):
        output = str(path_to_output) + '/%s' % question['TableIndex']
        result = '\n\n* Table %s.\n\n' % question['TableIndex']
        result += "OMS /SELECT TABLES\n    /IF SUBTYPES = ['Custom Table','Comparisons of Proportions']\n"
        result += '    /DESTINATION FORMAT = XLSX\n'
        result += "     OUTFILE = '%s'\n\n" % (output)
        return result

    def write_ctable(self, question, column_specs):
        result = "CTABLES\n  /FORMAT EMPTY=BLANK MISSING='.'\n  /SMISSING VARIABLE\n"
        result += "  /VLABELS VARIABLES=%s " % question['VariableName']
        result += 'Total '
        for spec in column_specs:
            if "$" in spec:
                result += '$C%s ' % spec.replace("$","")
                self.grouped_question.append(spec)
            else:
                result += 'C%s ' % spec
        result += 'DISPLAY=LABEL\n'
        result += "  /TABLE %s [C] BY Total [C][COUNT F40.0, COLPCT.COUNT PCT40.0] " % question['VariableName']
        for spec in column_specs:
            if spec in self.grouped_question:
                result += '+ $C%s [C] ' % spec.replace("$","")
                if self.__filtering_variable is None:
                    result += '[COUNT F40.0, COLPCT.COUNT PCT40.0] '
                else:
                    result += '> C%s [C] ' % self.__filtering_variable
            else:
                result += '+ C%s [C] ' % spec
                if self.__filtering_variable is None:
                    result += '[COUNT F40.0, COLPCT.COUNT PCT40.0] '
                else:
                    result += '> C%s [C] ' % self.__filtering_variable
                result += '[COUNT F40.0, COLPCT.COUNT PCT40.0] '
        result += '\n  /SLABELS POSITION=ROW VISIBLE=NO\n'
        result += "  /CATEGORIES VARIABLES=%s " % question['VariableName']
        result += "ORDER=A KEY=VALUE EMPTY=INCLUDE TOTAL=YES LABEL='Sigma' POSITION=AFTER\n"
        result += '  /CATEGORIES VARIABLES=Total '
        for spec in column_specs:
            if spec not in self.grouped_question:
                result += 'C%s ' % spec
        result += 'ORDER=A KEY=VALUE EMPTY=INCLUDE\n'
        result += '  /CATEGORIES VARIABLES=Total '
        for spec in column_specs:
            if spec in self.grouped_question:
                result += '$C%s ' % spec.replace("$","")
        result += 'ORDER=A KEY=VALUE EMPTY=INCLUDE\n'
        result += '  /CRITERIA CILEVEL=95\n'
        result += '  /TITLES\n'

        title = question['Title']
        title = title.replace('"', '')
        table_index = int(question['TableIndex'])
        if table_index < 10:
            result += "    TITLE='Table 0%s - %s: %s'\n" % (table_index, question['VariableName'], title)
        else:
            result += "    Title='Table %s - %s: %s'\n" % (table_index, question['VariableName'], title)
        if question['Base'] is not '':
            result += "    CORNER='%s - %s'\n" % ('Base', question['Base'])
        result += '  /COMPARETEST TYPE=PROP ALPHA=0.05 ADJUST=BONFERRONI ORIGIN=COLUMN INCLUDEMRSETS=YES\n'
        result += '    CATEGORIES=ALLVISIBLE MERGE=NO SHOWSIG=NO.\n\n'
        result += 'OMSEND.'
        return result

    def add_column_recode(self, column_specs, column_variables):
        result = ""
        for spec in column_specs:
            if "$" in spec:
                column = spec.replace("$","")
            else:
                column = spec
            if column in column_variables:
                result += "string C%s (A16).\n" % (column)
            result += "compute C%s = %s. \n" % (column, column)
            result += "APPLY DICTIONARY \n  /FROM *\n  /SOURCE VARIABLES="
            result += "%s\n  /TARGET VARIABLES=C%s\n" % (column, column)
            result += "  /FILEINFO\n  /VARINFO LEVEL VALLABELS=REPLACE VARLABEL."
            result += "\n  freq C%s %s.\n\n" % (column, column)
            result += "variable labels C%s '%s'.\n\n" % (column, column)
        result += "compute Total = 1.\n"
        return result
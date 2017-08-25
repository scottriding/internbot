import csv

class TableScript(object):
    
    def compile_scripts(self, path_to_tables, path_to_output):
        table_path = str(path_to_tables) + '/Tables to run.csv'
        output_path = str(path_to_output) + '/table script.txt'
        output = open(output_path, "w+")
        script = self.write_script(table_path, path_to_output)
        output.write(script)
                
    def write_script(self, path_to_tables, path_to_output):
        result = ''
        with open(path_to_tables, 'rb') as table_file:
            column_specs = self.compile_specs(path_to_tables)
            file = csv.DictReader(table_file, quotechar = '"')
            for table in file:
                result += self.write_table(table, path_to_output)
                result += self.write_ctable(table, column_specs)
        return result

    def compile_specs(self, path_to_tables):
        with open(path_to_tables, 'rb') as table_file:
            file = csv.DictReader(table_file, quotechar = '"')
            column_specs = []
            for row in file:
                if row['Column specs'] is not '':
                    column_specs.append(row['VariableName'])
        return column_specs

    def write_table(self, question, path_to_output):
        output = str(path_to_output) + '/%s' % question['TableIndex']
        result = '* Table %s.\n\n' % question['TableIndex']
        result += "OMS /SELECT TABLES\n    /IF SUBTYPES = ['Custom Table','Comparisons of Proportions']\n"
        result += '    /DESTINATION FORMAT = XLSX\n'
        result += "     OUTFILE = '%s'\n\n" % (output)
        return result

    def write_ctable(self, question, column_specs):
        result = '* Custom Tables.\nCTABLES\n'
        result += "  /FORMAT EMPTY=BLANK MISSING='.'\n  /SMISSING VARIABLE\n"
        result += "  /VLABELS VARIABLES=%s " % question['VariableName']
        result += 'Total '
        for spec in column_specs:
            result += 'C%s ' % spec
        result += 'DISPLAY=DEFAULT\n'
        result += "  /TABLE %s [C][COUNT F40.0, COLPCT.COUNT PCT40.0] BY Total [C] " % question['VariableName']
        for spec in column_specs:
            if '$' in spec:
                temp = spec.replace("$", "")
                result += '+ $C%s [C] ' %temp
            else:
                result+= '+ C%s [C] ' %spec
        result += '\n  /SLABELS POSITION=ROW VISIBLE=NO\n'
        result += "  /CATEGORIES VARIABLES=%s " % question['VariableName']
        result += 'ORDER=A KEY=VALUE EMPTY=INCLUDE TOTAL=YES POSITION=AFTER\n'
        result += '  /CATEGORIES VARIABLES=Total '
        for spec in column_specs:
            result += '%s ' %spec
        result += 'ORDER=A KEY=VALUE EMPTY=EXCLUDE\n'
        result += '  /CRITERIA CILEVEL=95\n'
        result += '  /TITLES\n'
        result += "    TITLE='Table %s - %s: %s'\n" % (question['TableIndex'], question['VariableName'], question['Title'])
        if question['Base'] is not '':
            result += "    CORNER='%s - %s'\n" % ('Base', question['Base'])
        result += '  /COMPARETEST TYPE=PROP ALPHA=0.05 ADJUST=BONFERRONI ORIGIN=COLUMN INCLUDEMRSETS=YES\n'
        result += '    CATEGORIES=ALLVISIBLE MERGE=NO SHOWSIG=NO.\n\n'
        result += 'OMSEND.'
        result += '\n\n\n'
        return result
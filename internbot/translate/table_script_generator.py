import csv

class TableScript(object):
    
    def compile_scripts(self, path_to_tables, path_to_output):
        table_path = str(path_to_tables) + '/Tables to run.csv'
        output_path = str(path_to_output) + '/table script.txt'
        output = open(output_path, "w+")
        script = self.write_script(table_path, output_path)
        output.write(script)
                
    def write_script(self, path_to_tables, path_to_output):
        
        result = ''
        with open(path_to_tables, 'rb') as table_file:
            file = csv.DictReader(table_file, quotechar = '"')
            for table in file:
                result += self.write_table(table, path_to_output)
                result += self.write_ctable(table)
        return result

    def write_table(self, question, path_to_output):
        output = str(path_to_output) + '/%s' % question['TableIndex']
        result = '* Table %s.\n\n' % question['TableIndex']
        result += "OMS /SELECT TABLES\n    /IF SUBTYPES = ['Custom Table','Comparisons of Proportions']\n"
        result += '    /DESTINATION FORMAT = XLSX\n'
        result += "     OUTFILE = '%s'\n\n" % (output)
        return result

    def write_ctable(self, question):
        result = '* Custom Tables.\nCTABLES\n'
        result += "  /FORMAT EMPTY=BLANK MISSING='.'\n  /SMISSING VARIABLE\n"
        result += "  /VLABELS VARIABLES=%s " % question['VariableName']
        result += 'Total Ccountry Csex Cage Cpatron DISPLAY=DEFAULT'
        result += '\n\n\n'
        return result
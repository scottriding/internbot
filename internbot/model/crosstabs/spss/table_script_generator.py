import csv

class TableScriptGenerator(object):

    def __init__(self):
        self.grouped_question = []

    def compile_scripts(self, tables, banners, embedded_variables, filtering_variable, path_to_output):        
        self.__tables = tables
        self.__banners = banners
        self.__embedded_variables = embedded_variables
        self.__filtering_variable = filtering_variable

        output = open(path_to_output, "w+")
        script = self.write_script(path_to_output)
        output.write(script)
                
    def write_script(self, path_to_output):
        result = ''
        result += self.banner_recodes()
        for table in self.__tables:
            result += self.write_table(table, path_to_output)
            result += self.write_ctable(table)
        return result

    def write_table(self, table, path_to_output):
        output = str(path_to_output) + '/%s' % table.index
        result = '\n\n* Table %s.\n\n' % table.index
        result += "OMS /SELECT TABLES\n    /IF SUBTYPES = ['Custom Table','Comparisons of Proportions']\n"
        result += '    /DESTINATION FORMAT = XLSX\n'
        result += "     OUTFILE = '%s'\n\n" % (output)
        return result

    def write_ctable(self, table):
        result = "CTABLES\n  /FORMAT EMPTY=BLANK MISSING='.'\n  /SMISSING VARIABLE\n"
        result += "  /VLABELS VARIABLES=%s " % table.name
        result += 'Total '
        for banner in self.__banners:
            if "$" in banner:
                result += '$C%s ' % banner.replace("$","")
                self.grouped_question.append(banner)
            else:
                result += 'C%s ' % banner
        result += 'DISPLAY=LABEL\n'
        result += "  /TABLE %s [C] BY Total [C][COUNT F40.0, COLPCT.COUNT PCT40.0] " % table.name
        for banner in self.__banners:
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
        result += "  /CATEGORIES VARIABLES=%s " % table.name
        result += "ORDER=A KEY=VALUE EMPTY=INCLUDE TOTAL=YES LABEL='Sigma' POSITION=AFTER\n"
        result += '  /CATEGORIES VARIABLES=Total '
        for banner in self.__banners:
            if spec not in self.grouped_question:
                result += 'C%s ' % spec
        result += 'ORDER=A KEY=VALUE EMPTY=INCLUDE\n'
        result += '  /CATEGORIES VARIABLES=Total '
        for banner in self.__banners:
            if banner in self.grouped_question:
                result += '$C%s ' % spec.replace("$","")
        result += 'ORDER=A KEY=VALUE EMPTY=INCLUDE\n'
        result += '  /CRITERIA CILEVEL=95\n'
        result += '  /TITLES\n'

        title = table.title
        title = title.replace('"', '')
        table_index = int(table.index)
        if table_index < 10:
            result += "    TITLE='Table 0%s - %s: %s'\n" % (table_index, table.name, title)
        else:
            result += "    Title='Table %s - %s: %s'\n" % (table_index, table.name, title)
        if table.base is not '':
            result += "    CORNER='%s - %s'\n" % ('Base', table.base)
        result += '  /COMPARETEST TYPE=PROP ALPHA=0.05 ADJUST=BONFERRONI ORIGIN=COLUMN INCLUDEMRSETS=YES\n'
        result += '    CATEGORIES=ALLVISIBLE MERGE=NO SHOWSIG=NO.\n\n'
        result += 'OMSEND.'
        return result

    def banner_recodes(self):
        result = ""
        for banner in self.__baners:
            if banner in self.__embedded_variables:
                result += "string C%s (A24).\n" % banner
            result += "compute C%s = %s. \n" % (banner, banner)
            result += "APPLY DICTIONARY \n  /FROM *\n  /SOURCE VARIABLES="
            result += "%s\n  /TARGET VARIABLES=C%s\n" % (banner, banner)
            result += "  /FILEINFO\n  /VARINFO LEVEL VALLABELS=REPLACE VARLABEL."
            result += "\n  freq C%s %s.\n\n" % (banner, banner)
            result += "variable labels C%s '%s'.\n\n" % (banner, banner)
        result += "compute Total = 1.\n"
        return result
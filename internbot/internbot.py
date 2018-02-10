import argparse
import base
import sys
import crosstabs

if __name__ == '__main__':

    parser = argparse.ArgumentParser(
        description='Automates Y2 Analytics reports.'
    )

    parser.add_argument('qsf', help='path to the Qualtrics QSF file')
    parser.add_argument('report', help='Report to be generated (either basic topline, full topline, appendix, powerpoint, and SPSS)')
    parser.add_argument('output', help='path to output folder')
    parser.add_argument('-xtabs', '--xlsx',help='path to crosstabs xlsx file')
    parser.add_argument('-tables','--table_csv',help='path to tables to run csv')

    args = parser.parse_args()

    if not args.report and not args.output:
        parser.print_help()
        sys.exit()

    compiler = base.QSFSurveyCompiler()
    survey = compiler.compile(args.qsf)
    
    if args.report == 'SPSS':    
        variables = crosstabs.Generate_Prelim_SPSS_Script.SPSSTranslator()
        tables = crosstabs.Generate_Prelim_SPSS_Script.TableDefiner()
        
        variables.define_variables(survey, args.output)
        tables.define_tables(survey, args.output)

    elif args.report == 'table_script':
        script = crosstabs.Generate_Table_Script.TableScript()
        script.compile_scripts(args.table_csv, args.output)

    elif args.report == 'final_touches':
        renamer = crosstabs.Polish_Final_Report.RenameTabs()
        renamer.rename(args.xlsx, args.table_csv, args.output)

        highlighter = crosstabs.Polish_Final_Report.Highlighter(args.output)
        highlighter.highlight(args.output)

    else:
        parser.print_help()
        sys.exit()


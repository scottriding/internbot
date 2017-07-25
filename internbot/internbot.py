import argparse
import base
import sys
import translate
import topline


if __name__ == '__main__':

    parser = argparse.ArgumentParser(
        description='Automates Y2 Analytics reports.'
    )

    parser.add_argument('qsf', help='path to the Qualtrics QSF file')
    parser.add_argument('report', help='Report to be generated (either topline, powerpoint, or SPSS)')
    parser.add_argument('output', help='path to output file')
    parser.add_argument('template', help='path to topline template')
    #parser.add_argument('csv',help='path to frequency csv')
    

    args = parser.parse_args()

    if not args.qsf and not args.report and not args.output:
        parser.print_help()
        sys.exit()

    compiler = base.QSFSurveyCompiler()
    survey = compiler.compile(args.qsf)
    report = topline.ReportGenerator(survey)
    
    if args.report == 'SPSS':
        translator = translate.SPSSTranslator()
        translator.define_variables(survey, args.output)
        
    if args.report == 'topline':
        report.generate_docx(args.csv, args.template, args.output)
        
    if args.report == 'powerpoint':
        report.generate_ppt(args.template, args.output)


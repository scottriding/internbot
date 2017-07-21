import argparse
import base
import sys
import translate
# import topline


if __name__ == '__main__':

    parser = argparse.ArgumentParser(
        description='Automates Y2 Analytics reports.'
    )

    parser.add_argument('qsf', help='path to the Qualtrics QSF file')
    parser.add_argument('report', help='Report to be generated (either topline or SPSS)')
    parser.add_argument('output', help='path to output file')

    args = parser.parse_args()

    if not args.qsf and not args.report and not args.output:
        parser.print_help()
        sys.exit()

    compiler = base.QSFSurveyCompiler()
    survey = compiler.compile(args.qsf)

    if args.report == 'SPSS':
        translater = translate.SPSSTranslator()
        translator.define_variables(survey, args.output)

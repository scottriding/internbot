import argparse
import base
import sys

if __name__ == '__main__':

    parser = argparse.ArgumentParser(
        description='Automates Scoring Microsoft Assessment Surveys'
    )

    parser.add_argument('qsf', help='path to the Qualtrics QSF file')
    parser.add_argument('csv', help='path to csv with questions')
    parser.add_argument('output', help='path to output folder')
    
    args = parser.parse_args()

    if not args.qsf and not args.report and not args.output:
        parser.print_help()
        sys.exit()

    compiler = base.QSFSurveyCompiler()
    scores = compiler.grab_scoring(args.qsf)
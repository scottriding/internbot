import argparse
import base
import sys
import topline
import data_analysis

if __name__ == '__main__':

    parser = argparse.ArgumentParser
    (
        description='Automates Y2 Analytics reports.'
    )

    parser.add_argument('qsf', help='path to the Qualtrics QSF file')
    parser.add_argument('report', help='Report to be generated (either basic topline, full topline, appendix, powerpoint, and SPSS)')
    parser.add_argument('output', help='path to output folder')
    parser.add_argument('-temp','--template', help='path to topline/powerpoint/appendix template')
    parser.add_argument('-freq','--freq_csv', help='path to frequency csv')
    parser.add_argument('-app','--oe_csv', help='path to open ended responses csv')

    args = parser.parse_args()

    if not args.qsf and not args.report and not args.output:
        parser.print_help()
        sys.exit()

    compiler = base.QSFSurveyCompiler()
    survey = compiler.compile(args.qsf)
    report = topline.ReportGenerator(survey)

    elif args.report == 'graphs':
        translator = data_analysis.GraphDefiner()
        translator.define_graphs(survey, args.output)
        
    elif args.report == 'basic_topline':
        report.generate_basic_topline(args.freq_csv, args.template, args.output)

    elif args.report == 'full_topline':
        report.generate_full_topline(args.freq_csv, args.template, args.output, args.oe_csv)

    elif args.report == 'appendix':
        report.generate_appendix(args.template, args.oe_csv, args.output)
        
    elif args.report == 'powerpoint':
        report.generate_ppt(args.template, args.output)

    else:
        parser.print_help()
        sys.exit()
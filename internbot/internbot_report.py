import argparse
import base
import sys
import topline
import data_analysis

if __name__ == '__main__':

    parser = argparse.ArgumentParser(description='Automates Y2 Analytics topline reports.')

    parser.add_argument('report', help='Report to be generated: basic_topline/full_topline/appendix/powerpoint/graphs.')
    parser.add_argument('output', help='Path to output folder.')
    parser.add_argument('template', help='Path to topline/powerpoint/appendix templates.')
    parser.add_argument('-qsf','--qsf', help='Path to Qualtrics QSF file.')
    parser.add_argument('-csv','--csv', help='Path to question/frequencies csv.')
    parser.add_argument('-app','--oe_csv', help='Path to open ended responses csv.')

    args = parser.parse_args()

    if not args.report and not args.temp and not args.output:
        parser.print_help()
        sys.exit()

    isQSF = False

    if args.qsf:
        compiler = base.QSFSurveyCompiler()
        survey = compiler.compile(args.qsf)
        report = topline.QSF.ReportGenerator(survey)
        isQSF = True

    elif args.csv:
        report = topline.CSV.ReportGenerator(args.csv)

    if args.report == 'graphs':
        translator = data_analysis.GraphDefiner()
        translator.define_graphs(survey, args.output)

    elif args.report = 'automate_MS':
        pass
        
    elif args.report == 'basic_topline':
        if isQSF is True:
            report.generate_basic_topline(args.csv, args.template, args.output)
        else:
            report.generate_basic_topline(args.template, args.output)

    elif args.report == 'full_topline':
        report.generate_full_topline(args.csv, args.template, args.output, args.oe_csv)

    elif args.report == 'appendix':
        report.generate_appendix(args.template, args.oe_csv, args.output)
        
    elif args.report == 'powerpoint':
        report.generate_ppt(args.template, args.output)

    else:
        parser.print_help()
        sys.exit()
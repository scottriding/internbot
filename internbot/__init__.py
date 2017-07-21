print('I do everything except coffee runs.')

import base
import topline

path_to_qsf = '/Users/y2analytics/Downloads/Weber_School_District_Potential_Bond_Viability_Survey.qsf'
path_to_csv = '/Users/y2analytics/Downloads/Weber School District Weighted Freqs TMF v2.csv'
path_to_template = '/Users/y2analytics/Dropbox (Y2 Analytics)/Y2 Analytics Team Folder/R&D/Topline automation/Template/topline_template.docx'
path_to_output = '/Users/y2analytics/Desktop/Weber School District topline.docx'

compile = base.QSFSurveyCompiler()
survey = compile.compile(path_to_qsf)
generator = topline.ReportGenerator(survey)
generator.generate_report(path_to_csv, path_to_template, path_to_output)

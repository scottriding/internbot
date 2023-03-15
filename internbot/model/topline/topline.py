from model.topline import appendix
from model.topline import document
from model.topline import powerpoint

from model.topline import assigner

class Topline(object):

	def build_appendix_model(self, path_to_csv):
		appendix.Appendix().check_input(path_to_csv)
		frequency_assigner = assigner.Assigner()
		frequency_assigner.build_questions(path_to_csv)
		return frequency_assigner.assign()

	def build_appendix_report(self, question_blocks, path_to_output, template_path):
		appendix.Appendix().build_appendix_report(question_blocks, path_to_output, template_path)

	def build_document_model(self, path_to_csv, groups, survey):
		assigner.Assigner().check_input(path_to_csv, groups)
		frequency_assigner = assigner.Assigner()
		frequency_assigner.build_questions(path_to_csv, groups, survey)
		return frequency_assigner.assign()

	def build_document_report(self, question_blocks, path_to_template, path_to_output):
		document.Document().build_document_report(question_blocks, path_to_template, path_to_output)
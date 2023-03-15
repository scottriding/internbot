from model.base import qsf_parser
import json

class QSFSurveyCompiler(object):

	def __init__(self):
		self.survey_parser = qsf_parser.QSFSurveyParser()
		self.blockflow_parser = qsf_parser.QSFBlockFlowParser()
		self.blocks_parser = qsf_parser.QSFBlocksParser()
		self.questions_parser = qsf_parser.QSFQuestionsParser()
		self.scoring_parser = qsf_parser.QSFScoringParser()

	def compile(self, path_to_qsf):
		if path_to_qsf is not None:
			try:
				qsf_json = self.parse_json(path_to_qsf)
				survey = self.compile_survey(qsf_json)
				return survey
			except Exception as e:
				raise ValueError(f'Could not parse survey file. This is not a user error, file may be corrupted. \nPlease contact Systems Developer.')
		else:
			return None

	def grab_scoring(self, path_to_qsf):
		qsf_json = self.parse_json(path_to_qsf)
		scoring = self.compile_scores(qsf_json)
		return scoring

	def parse_json(self, path_to_qsf):
		with open(path_to_qsf) as file:
			qsf_file = json.load(file)
		return qsf_file

	def compile_survey(self, qsf_json):
		survey = self.survey_parser.parse(qsf_json['SurveyEntry'])
		blocks = self.compile_blocks(qsf_json)
		questions = self.compile_questions(qsf_json)
		for block in blocks:
			survey.add_block(block)
		survey.add_questions(questions)
		return survey

	def compile_blocks(self, qsf_json):
		block_ids = self.blockflow_parser.parse(self.find_element('FL', qsf_json))
		blocks = self.blocks_parser.parse(self.find_element('BL', qsf_json))
		blocks.sort(block_ids)
		return blocks

	def compile_questions(self, qsf_json):
		return self.questions_parser.parse(self.find_elements('SQ', qsf_json))

	def compile_scores(self, qsf_json):
		return self.scoring_parser.parse(self.find_elements('SCO', qsf_json))

	def find_element(self, element_name, qsf_json):
		return next(iter(self.find_elements(element_name, qsf_json)), None)

	def find_elements(self, element_name, qsf_json):
		elements = qsf_json['SurveyElements']
		return [element for element in elements if element['Element'] == element_name]


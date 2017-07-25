from pptx import Presentation
from pptx.util import Inches, Pt

class ToplinePPT(object):
    
    def __init__ (self, questions, path_to_template):
        self.ppt = Presentation(path_to_template)
        self.questions = questions

    def save(self, path_to_output):
        self.write_questions()
        self.save_file(path_to_output)

    def write_questions(self):
        for question in self.questions:
            self.write_question(question)

    def save_file(self, path_to_output):
        self.ppt.save(path_to_output)

    def write_question(self, question):
        slide = self.ppt.slides.add_slide(self.ppt.slide_layouts[1])
        self.write_name(question.name, slide)
        self.write_prompt(question.prompt, slide)

    def write_name(self, name, slide):
        left = top = width = height = Inches(1)
        txBox = slide.shapes.add_textbox(left, top, width, height)
        tf = txBox.text_frame
        tf.text = name

    def write_prompt(self, prompt, slide):
        left = top = width = height = Inches(2)
        txBox = slide.shapes.add_textbox(left, top, width, height)
        tf = txBox.text_frame
        tf.text = prompt

    def get_ppt(self):
        return self.ppt
        
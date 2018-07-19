from docx import Document
from docx.shared import Inches
   
class ToplineReport(object):

    def __init__ (self, questions, path_to_template):
        self.doc = Document(path_to_template)
        self.line_break = self.doc.styles['LineBreak']
        self.questions = questions
        names = questions.list_names()
        for name in names:
            self.write_question(name)
        
    def write_question(self, name):
        if self.questions.display_logic(name) != "":
            display_prompt = self.doc.add_paragraph()
            display_prompt.add_run(self.questions.display_logic(name))
            self.doc.add_paragraph()
        paragraph = self.doc.add_paragraph() # each question starts a new paragraph
        self.write_name(name, paragraph)
        self.write_prompt(self.questions.prompt(name), paragraph)
        self.write_responses(self.questions.return_responses(name))
        new = self.doc.add_paragraph("") # space between questions
        new.style = self.line_break
        self.doc.add_paragraph("") # space between questions
        
    def write_name(self, name, paragraph):
        paragraph.add_run(name + ".")
    
    def write_prompt(self, prompt, paragraph):
        print prompt
        paragraph_format = paragraph.paragraph_format
        paragraph_format.keep_together = True           # question prompt will all be fit in one page
        paragraph_format.left_indent = Inches(1)        
        paragraph.add_run("\t" + prompt)                
        paragraph_format.first_line_indent = Inches(-1) # hanging indent if necessary
        
    def write_n(self, n, paragraph):
        paragraph.add_run("(n = " + str(n) + ")")
    
    def write_responses(self, responses):
        table = self.doc.add_table(rows = 1, cols = 5)
        first_row = True
        for response in responses:
            response_cells = table.add_row().cells
            response_cells[1].merge(response_cells[2])
            response_cells[1].text = response
            if first_row == True:
                text = self.freqs_percent(responses[response])
                if text != "*":           
                    response_cells[3].text = text + "%"
                    first_row = False   
                else:
                    response_cells[3].text = text                    
            else:
                response_cells[3].text = self.freqs_percent(responses[response])
        
    def save(self, path_to_output):
        self.doc.save(path_to_output)
        
    def freqs_percent(self, freq):
        percent = freq * 100
        if percent > 0 and percent < 1:
            return "<1"
        if percent == 0:
            return "*"
        result = int(round(percent))
        return str(result)
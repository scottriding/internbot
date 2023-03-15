import docx
from docx.shared import Inches
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

from collections import OrderedDict

class Document(object):

    def build_document_report(self, question_blocks, path_to_template, path_to_output):
        self.__question_blocks = question_blocks
        self.__document = docx.Document(path_to_template)

        # depending on template, this could not be in styles
        try:
            self.line_break = self.__document.styles['LineBreak']
        except KeyError:
            self.line_break = None

        self.write_questions()
        self.__document.save(path_to_output)

    def write_questions(self):
        for question in self.__question_blocks.questions:
            if question.parent == 'CompositeQuestion':
                if question.type == "CompositeMatrix":
                    self.write_matrix_question(question)
                else:
                    self.write_binary_question(question)
            elif question.type == 'TE':
                self.write_open_ended(question)
            elif question.type == 'DB':
                self.write_descriptive(question)
            else:
                self.write_question(question)

            line_break = self.__document.add_paragraph("")
            if self.line_break is not None:
                line_break.style = self.line_break
            self.__document.add_paragraph("")

    def write_matrix_question(self, matrix_question):
        ## if any of the sub_questions are grouped, break them all up
        break_up = False
        for sub_question in matrix_question.questions:
            total_ns = self.pull_ns(sub_question)
            if len(total_ns) > 1:
                break_up = True

        if break_up:
            paragraph = self.__document.add_paragraph()
            paragraph_format = paragraph.paragraph_format
            paragraph_format.keep_together = True
            paragraph.add_run(matrix_question.prompt + ". ")

            self.__document.add_paragraph("")
            self.__document.add_paragraph("")

            for sub_question in matrix_question.questions:
                self.write_question(sub_question)
                self.__document.add_paragraph("")
                self.__document.add_paragraph("")
        else:
            paragraph = self.__document.add_paragraph()
            self.write_name(matrix_question.name, paragraph)
            self.write_prompt(matrix_question.prompt, paragraph)

            table = self.__document.add_table(rows = 1, cols = 0)
            table.add_column(width = Inches(1))
            table.add_column(width = Inches(1))
            first_row = True
            header_cells = table.add_row().cells

            for sub_question in matrix_question.questions:
                if first_row:
                    for response in sub_question.responses:
                        response_cells = table.add_column(width = Inches(1)).cells
                        response_cells[1].text = response.label
                    first_row = False

                total_ns = self.pull_ns(sub_question)
                self.group_sub_matrix_question(sub_question, total_ns, table, first_row)

    def group_sub_matrix_question(self, sub_question, total_ns, table, first_row):
        question_cells = table.add_row().cells
        n_text = ""
        header = "Basic"
        for group, n in total_ns.items():
            header = group
            n_text = "(n = %s)" % n
        question_cells[1].text = "%s %s" % (sub_question.prompt, n_text)
        index = 2
        first = True
        for response in sub_question.responses:
            for group, frequency in response.frequencies.frequencies.items():
                if group == header:
                    if frequency.stat == 'percent':
                        freq = self.percent(frequency.result, first_row)
                    elif frequency.stat == 'mean':
                        freq = self.mean(frequency.result)
                    else:
                        freq = str(frequency.result)
                if first:
                    percent = str(freq) + "%"
                    question_cells[index].text = percent
                    first = False
                else:
                    question_cells[index].text = freq
                index += 1

    def write_binary_question(self, binary_question):
        paragraph = self.__document.add_paragraph()
        self.write_name(binary_question.name, paragraph)
        self.write_prompt(binary_question.prompt, paragraph)

        groups = self.pull_binary_groups(binary_question)
        if len(groups) == 1:
            table = self.__document.add_table(rows=1, cols=5)
            first_row = True
            cells = table.add_row().cells
            cells[1].merge(cells[2])
            cells[3].text = "Average"
            for sub_question in binary_question.questions:
                cells = table.add_row().cells
                cells[1].merge(cells[2])
                response = next((response for response in sub_question.responses if response.value is not None), None)
                if response is not None:
                    if not response.frequencies:
                        shading_elm = parse_xml(r'<w:shd {} w:fill="FFF206"/>'.format(nsdecls('w')))
                        cells[1]._tc.get_or_add_tcPr().append(shading_elm)
                    else:
                        for group, frequency in response.frequencies.frequencies.items():
                            cells[1].text = "%s (n = %s)" % (sub_question.prompt, frequency.population)
                            if frequency.stat == 'percent':
                                freq = self.percent(frequency.result, first_row)
                            elif frequency.stat == 'mean':
                                freq = self.mean(frequency.result)
                            else:
                                freq = str(frequency.result)
                            cells[3].text = freq
                    first_row = False
                
        else:
            table = self.__document.add_table(rows=1, cols=len(groups)+4)
            titles_row = table.add_row().cells
            titles_row[1].merge(titles_row[2])
            headers_index = 0
            first_row = True
            first_group = True
            row = -1
            table_rows = []
            while headers_index < len(groups):
                header_text = "Average %s" % groups[headers_index]
                titles_row[headers_index+4].text = header_text
                for sub_question in binary_question.questions:
                    if first_group:
                        cells = table.add_row().cells
                        cells[1].merge(cells[2])
                        cells[1].text = sub_question.prompt
                        table_rows.append(cells)
                        row += 1
                    else:
                        row += 1
                    for response in sub_question.responses:
                        if not response.frequencies:
                            shading_elm = parse_xml(r'<w:shd {} w:fill="FFF206"/>'.format(nsdecls('w')))
                            table_rows[row][1]._tc.get_or_add_tcPr().append(shading_elm)
                        else:
                            for group, frequency in response.frequencies.frequencies.items():
                                if group == groups[headers_index]:
                                    if frequency.stat == 'percent':
                                        freq = self.percent(frequency.result, first_row)
                                    elif frequency.stat == 'mean':
                                        freq = self.mean(frequency.result)
                                    else:
                                        freq = str(frequency.result)
                                    table_rows[row][headers_index+4].text = freq
                                    n_text = " (%s n = %s)" % (group, frequency.population)
                                    table_rows[row][1].text += n_text
                                first_row = False
                headers_index += 1
                row = -1
                first_group = False

    def write_open_ended(self, question):
        paragraph = self.__document.add_paragraph()  # each question starts a new paragraph
        self.write_name(question.name, paragraph)
        self.write_prompt(question.prompt, paragraph)
        paragraph.add_run(' (OPEN-ENDED RESPONSES VERBATIM IN APPENDIX)')

    def write_descriptive(self, question):
        paragraph = self.__document.add_paragraph()  # each question starts a new paragraph
        self.write_name(question.name, paragraph)
        self.write_prompt(question.prompt, paragraph)

    def write_question(self, question):
        paragraph = self.__document.add_paragraph()
        self.write_name(question.name, paragraph)
        self.write_prompt(question.prompt, paragraph)
        self.write_responses(question, paragraph)

    def write_name(self, name, paragraph):
        paragraph.add_run(name + ".")

    def write_prompt(self, prompt, paragraph):
        paragraph_format = paragraph.paragraph_format
        paragraph_format.keep_together = True
        paragraph_format.left_indent = Inches(1)
        paragraph.add_run("\t" + prompt + " ")
        paragraph_format.first_line_indent = Inches(-1)

    def write_responses(self, question, paragraph):
        total_ns = self.pull_ns(question)
        if len(total_ns) == 1:
            ## responses only belong to one group
            for group, n in total_ns.items():
                n_text = "(n = %s)" % n
                paragraph.add_run(n_text)
            self.write_results(question.responses, total_ns)
        else:
            self.write_grouped_results(question.responses, total_ns)

    def pull_ns(self, question):
        ## we're going to total all the groups that are in the responses of this question
        total_ns = OrderedDict()
        for response in question.responses:
            for group, frequency in response.frequencies.frequencies.items():
                if frequency.result != "NA":
                ### if the response belongs to a group, calculate the n
                    if total_ns.get(group) is not None:
                        total_ns[group] += int(frequency.population)
                    else:
                        total_ns[group] = int(frequency.population)
        return total_ns

    def pull_binary_groups(self, binary_question):
        ## we're going to total all the groups that are in the responses of this question
        total_groups = []
        for sub_question in binary_question.questions:
            for response in sub_question.responses:
                for group, frequency in response.frequencies.frequencies.items():
                    if frequency.result != "NA":
                        if group not in total_groups:
                            total_groups.append(group)
        return total_groups

    def write_results(self, responses, total_ns):
        header = list(total_ns.keys())[0]
        table = self.__document.add_table(rows=1, cols=5)
        first_row = True
        for response in responses:
            response_cells = table.add_row().cells
            response_cells[1].merge(response_cells[2])
            response_cells[1].text = response.label
            if not response.frequencies:
                ## highlight the missing frequency
                shading_elm = parse_xml(r'<w:shd {} w:fill="FFF206"/>'.format(nsdecls('w')))
                response_cells[1]._tc.get_or_add_tcPr().append(shading_elm)
            for group, frequency in response.frequencies.frequencies.items():
                if group == header:
                    if frequency.stat == 'percent':
                        freq = self.percent(frequency.result, first_row)
                    elif frequency.stat == 'mean':
                        freq = self.mean(frequency.result)
                    else:
                        freq = str(frequency.result)
                    response_cells[3].text = freq
            first_row = False

    def write_grouped_results(self, responses, total_ns):
        headers = list(total_ns.keys())
        table = self.__document.add_table(rows=1, cols=len(headers) + 4)
        titles_row = table.add_row().cells
        titles_row[1].merge(titles_row[2])
        for header_index in range(len(headers)):
            header_text = "Total %s\n(n = %s)" % (headers[header_index], total_ns[headers[header_index]])
            titles_row[header_index + 4].text = header_text
        first_row = True
        for response in responses:
            response_cells = table.add_row().cells
            response_cells[1].merge(response_cells[3])
            response_cells[1].text = response.label
            freq_col = 4
            if not response.frequencies:
                ## highlight the missing frequency
                shading_elm = parse_xml(r'<w:shd {} w:fill="FFF206"/>'.format(nsdecls('w')))
                response_cells[1]._tc.get_or_add_tcPr().append(shading_elm)
            for header in headers:
                if response.frequencies.frequencies.get(header) is not None:
                    freq = response.frequencies.frequencies.get(header)
                    if freq.stat == 'percent':
                        text = self.percent(freq.result, first_row)
                    elif freq.stat == 'mean':
                        text = self.mean(freq.result)
                    else:
                        text = str(freq.result)
                    response_cells[freq_col].text = text
                first_row = False
                freq_col += 1           

    def percent(self, freq, is_first):
        try:
            if float(freq) >= 1.0:
                result = int(freq) * 100
            else:
                percent = float(freq)
                percent = percent * 100
                if percent >= 0 and percent < 1:
                    result = "<1"
                else:
                    result = int(round(percent))
            if is_first:
                result = str(result) + "%"
            return str(result)
        except:
            return freq

    def mean(self, freq):
        try:
            freq = float(freq)
            if freq >= 0 and freq < 1:
                result =  '<1'
            else:
                result = int(round(freq))
            return str(result)
        except:
            return freq
        
    def save(self, path_to_output):
        self.__document.save(path_to_output)
        


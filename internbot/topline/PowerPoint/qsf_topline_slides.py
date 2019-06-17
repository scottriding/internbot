
from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION, XL_LABEL_POSITION
from pptx.enum.dml import MSO_COLOR_TYPE, MSO_THEME_COLOR
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
class QSFToplineSlides(object):

    def __init__ (self, survey, path_to_template, path_to_output, years):
        self.presentation = Presentation(path_to_template)
        self.questions = survey.get_questions()
        self.survey = survey
        self.years = years
        self.path_to_output = path_to_output
        self.slide_layouts = self.presentation.slide_layouts

    def save(self, path_to_output):
        self.chart_questions()
        self.presentation.save(path_to_output)
        print("Finished!")

    def chart_questions(self):
        title_slide = self.presentation.slides[0]
        title_slide.shapes[0].text = self.survey.name
        title_slide.shapes[1].text = "Y2 Analytics Slide Automation"
        for question in self.questions:
            to_print = "Writing question: %s" % question.name
            print(to_print)
            if question.parent == 'CompositeQuestion':
                self.chart_composite_question(question)
            elif question.type == 'TE':
                pass
            else:
                self.chart_question(question)

    def chart_question(self, question):
        slide = self.presentation.slides.add_slide(self.presentation.slide_layouts.get_by_name('AutomatedChart'))
        shapes = slide.shapes
        self.write_question_details(slide, question)
        if question.type == 'RO':
            pass
        else:
            self.bar_chart_vertical(question, slide)
            self.bar_chart_horizontal(question, slide)
            self.stacked_bar_horizontal(question, slide)
            self.pie_chart(question, slide)

    def chart_composite_question(self, question):
        slide = self.presentation.slides.add_slide(self.presentation.slide_layouts.get_by_name('AutomatedChart'))
        shapes = slide.shapes
        self.write_question_details(slide, question)

        # if question.type == 'CompositeMatrix':
        #     self.write_matrix(question.questions)
        # elif question.type == 'CompositeConstantSum':
        #     self.write_allocate(question.questions)
        # else:
        #     self.write_binary(question.questions)

    # def write_matrix(self, sub_questions):
    #
    #     for sub_question in sub_questions:
    #
    #         for response in sub_question.responses:
    #             if response.has_frequency is True:
    #                 for year, frequency in response.frequencies.items():
    #                     question_cells[index].text = frequency
    #             else:
    #                 if first_row is True:
    #                     question_cells[index].text = "--%"
    #                 else:
    #                     question_cells[index].text = "--"
    #             first_row = False
    #             index += 1

    def write_question_details(self, slide, question):
        header = slide.shapes[0]
        header.text = str(question.name)
        sub_header = slide.shapes[1]
        sub_header.text = "Analytics Description"
        n_box = slide.shapes[2]
        n_box.text = "N=" +str(question.n)

    def bar_chart_vertical(self, question, slide):
        categories = []
        frequencies = []
        for response in question.responses:
            categories.append(str(response.response))
            for year, frequency in response.frequencies.items():
                frequencies.append(frequency)
        chart_data = CategoryChartData()
        chart_data.categories = categories
        chart_data.add_series('Series 1', iter(frequencies))
        left, top, width, height = Inches(0), Inches(1.5), Inches(6), Inches(2.5)

        chart = slide.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, left, top, width, height, chart_data).chart
        chart.chart_title.has_text_frame = True
        chart.chart_title.text_frame.text = "Q: " + str(question.prompt)
        chart.chart_title.text_frame.paragraphs[0].font.size = Pt(12)
        chart.has_legend = False
        category_axis = chart.category_axis
        category_axis.has_minor_gridlines = False
        category_axis.has_major_gridlines = False
        category_axis.tick_labels.font.size = Pt(12)
        value_axis = chart.value_axis
        value_axis.has_minor_gridlines = False
        value_axis.has_major_gridlines = False
        value_axis.tick_labels.font.size = Pt(12)
        chart.plots[0].has_data_labels = True
        data_labels = chart.plots[0].data_labels
        data_labels.number_format = '0%'
        data_labels.position = XL_LABEL_POSITION.OUTSIDE_END
        data_labels.font.size = Pt(14)
        plot = chart.plots[0]
        series = plot.series[0]
        fill = series.format.fill
        fill.solid()
        fill.fore_color.theme_color = MSO_THEME_COLOR.ACCENT_1


    def bar_chart_horizontal(self, question, slide):
        categories = []
        frequencies = []
        for response in question.responses:
            categories.append(str(response.response))
            for year, frequency in response.frequencies.items():
                frequencies.append(frequency)
        chart_data = CategoryChartData()
        chart_data.categories = categories
        chart_data.add_series('Series 1', iter(frequencies))
        left, top, width, height = Inches(6), Inches(1.5), Inches(6), Inches(2.5)

        chart = slide.shapes.add_chart(XL_CHART_TYPE.BAR_CLUSTERED, left, top, width, height, chart_data).chart
        chart.chart_title.has_text_frame = True
        chart.chart_title.text_frame.text = "Q: " + str(question.prompt)
        chart.chart_title.text_frame.paragraphs[0].font.size = Pt(12)
        chart.has_legend = False
        category_axis = chart.category_axis
        category_axis.has_minor_gridlines = False
        category_axis.has_major_gridlines = False
        category_axis.tick_labels.font.size = Pt(12)
        value_axis = chart.value_axis
        value_axis.has_minor_gridlines = False
        value_axis.has_major_gridlines = False
        value_axis.tick_labels.font.size = Pt(12)
        chart.plots[0].has_data_labels = True
        data_labels = chart.plots[0].data_labels
        data_labels.number_format = '0%'
        data_labels.position = XL_LABEL_POSITION.OUTSIDE_END
        data_labels.font.size = Pt(14)
        plot = chart.plots[0]
        series = plot.series[0]
        fill = series.format.fill
        fill.solid()
        fill.fore_color.theme_color = MSO_THEME_COLOR.ACCENT_2


    def stacked_bar_vertical(self, question, slide):
        pass

    def stacked_bar_horizontal(self, question, slide):
        series_names = []
        frequencies = []
        for response in question.responses:
            series_names.append(str(response.response))
            for year, frequency in response.frequencies.items():
                frequencies.append(frequency)
        chart_data = CategoryChartData()
        chart_data.categories = ["Variable 1"]
        for idx in range(0, len(series_names)):
            temp = []
            temp.append(frequencies[idx])
            series = str(series_names[idx])
            chart_data.add_series(series, iter(temp))

        left, top, width, height = Inches(0), Inches(4.5), Inches(6), Inches(2.5)

        chart = slide.shapes.add_chart(XL_CHART_TYPE.BAR_STACKED, left, top, width, height, chart_data).chart
        chart.chart_title.has_text_frame = True
        chart.chart_title.text_frame.text = "Q: " + str(question.prompt)
        chart.chart_title.text_frame.paragraphs[0].font.size = Pt(12)
        chart.has_legend = True
        category_axis = chart.category_axis
        category_axis.has_minor_gridlines = False
        category_axis.has_major_gridlines = False
        category_axis.tick_labels.font.size = Pt(12)
        value_axis = chart.value_axis
        value_axis.has_minor_gridlines = False
        value_axis.has_major_gridlines = False
        value_axis.visible = False

        chart.legend.position = XL_LEGEND_POSITION.BOTTOM
        chart.legend.font.size = Pt(12)
        chart.legend.include_in_layout = False
        chart.plots[0].has_data_labels = True
        data_labels = chart.plots[0].data_labels
        data_labels.number_format = '0%'
        data_labels.font.size = Pt(12)




    def pie_chart(self, question, slide):
        categories = []
        frequencies = []
        for response in question.responses:
            categories.append(str(response.response))
            for year, frequency in response.frequencies.items():
                frequencies.append(frequency)
        chart_data = CategoryChartData()
        chart_data.categories = categories
        chart_data.add_series('Series 1', iter(frequencies))
        left, top, width, height = Inches(6), Inches(4.5), Inches(6), Inches(2.5)

        chart = slide.shapes.add_chart(XL_CHART_TYPE.PIE, left, top, width, height, chart_data).chart
        chart.chart_title.has_text_frame = True
        chart.chart_title.text_frame.text = "Q: " + str(question.prompt)
        chart.chart_title.text_frame.paragraphs[0].font.size = Pt(12)
        chart.has_legend = True
        # category_axis = chart.category_axis
        # category_axis.has_minor_gridlines = False
        # category_axis.has_major_gridlines = False
        # category_axis.tick_labels.font.size = Pt(12)
        # value_axis = chart.value_axis
        # value_axis.has_minor_gridlines = False
        # value_axis.has_major_gridlines = False
        # value_axis.tick_labels.font.size = Pt(12)
        chart.legend.position = XL_LEGEND_POSITION.BOTTOM
        chart.legend.font.size = Pt(12)
        chart.legend.include_in_layout = False
        chart.plots[0].has_data_labels = True
        data_labels = chart.plots[0].data_labels
        data_labels.number_format = '0%'
        data_labels.font.size = Pt(14)



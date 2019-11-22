from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION, XL_LABEL_POSITION
from pptx.enum.dml import MSO_COLOR_TYPE, MSO_THEME_COLOR
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor

class Powerpoint(object):

    def chart_questions(self, survey, path_to_template, path_to_output, years):
        self.presentation = Presentation(path_to_template)
        self.questions = survey.get_questions()
        self.survey = survey
        self.years = years
        self.path_to_output = path_to_output
        self.slide_layouts = self.presentation.slide_layouts

        title_slide = self.presentation.slides[0]
        title_slide.shapes[0].text = self.survey.name
        title_slide.shapes[1].text = "Y2 Analytics Slide Automation"
        for question in self.questions:
            to_print = "Writing question: %s" % question.name
            print(to_print)
            try:
                if question.parent == 'CompositeQuestion':
                    self.chart_composite_question(question)
                elif question.type == 'TE':
                    pass
                else:
                    self.chart_question(question)
            except ValueError:
                pass

    def save(self, path_to_output):
        self.chart_questions()
        self.presentation.save(path_to_output)
        print("Finished!")

    def chart_question(self, question):
        slide = self.presentation.slides.add_slide(self.presentation.slide_layouts.get_by_name('AutomatedChart'))
        shapes = slide.shapes
        self.write_question_details(slide, question)

        if len(self.years) > 0:
            self.column_clustered_basic_trended(question, slide)
            self.bar_clustered_basic_trended(question, slide)
        else:
            if len(question.responses) < 3:
                self.pie_simple(question, slide)
                self.column_clustered_basic(question, slide)
            else:
                self.column_clustered_basic(question, slide)
                self.bar_clustered_basic(question, slide)

    def chart_composite_question(self, question):
        if len(self.years) > 0:
            for sub_question in question.questions:
                self.chart_question(sub_question)
        else:
            if question.type == 'CompositeMatrix':
                slide_1 = self.presentation.slides.add_slide(self.presentation.slide_layouts.get_by_name('AutomatedChart'))
                self.write_composite_question_details(slide_1, question, 1, 3)
                self.column_clustered_matrix_by_categories(question, slide_1)
                self.column_clustered_matrix_by_scale_points(question, slide_1)

                slide_2 = self.presentation.slides.add_slide(self.presentation.slide_layouts.get_by_name('AutomatedChart'))
                self.write_composite_question_details(slide_2, question, 2, 3)
                self.bar_clustered_matrix_by_categories(question, slide_2)
                self.bar_clustered_matrix_by_scale_points(question, slide_2)

                slide_3 = self.presentation.slides.add_slide(self.presentation.slide_layouts.get_by_name('AutomatedChart'))
                self.write_composite_question_details(slide_3, question, 3, 3)
                self.stacked_bar_matrix(question, slide_3, 'top')
                self.stacked_bar_matrix(question, slide_3, 'bottom')

            elif question.type == 'CompositeConstantSum':
                slide_1 = self.presentation.slides.add_slide(self.presentation.slide_layouts.get_by_name('AutomatedChart'))
                self.write_composite_question_details(slide_1, question, 1, 1)
                self.column_clustered_allocation(question, slide_1)
                self.bar_clustered_allocation(question, slide_1)
            else:
                # self.write_binary(question.questions)
                pass

    def write_question_details(self, slide, question):
        header = slide.shapes[0]
        header.text = str(question.name)
        n_box = slide.shapes[1]
        #n_box.text = "N=" + str(question.n)
        count_box = slide.shapes[2]
        count_box.text = "Slide 1 of 1"

    def write_composite_question_details(self, slide, question, slide_number, slide_total):
        header = slide.shapes[0]
        header.text = str(question.name)
        n_box = slide.shapes[1]
        #n_box.text = "n's in charts"
        count_box = slide.shapes[2]
        count_box.text = "Slide "+str(slide_number)+" of "+str(slide_total)

    def max_years(self, responses):
        years_used = []
        for response in responses:
            for year in self.years:
                if response.frequencies.get(year) is not None:
                    if year not in years_used:
                        years_used.append(year)
        return years_used

    def max_years_subquestions(self, sub_questions):
        years_used = []
        for question in sub_questions:
            for response in question.responses:
                for year in self.years:
                    if response.frequencies.get(year) is not None:
                        if year not in years_used:
                            years_used.append(year)
        return years_used

    def column_clustered_basic(self, question, slide):
        categories = []
        frequencies = []
        for response in question.responses:
            categories.append(str(response.response))
            for year, frequency in response.frequencies.items():
                frequencies.append(frequency)
        chart_data = CategoryChartData()
        chart_data.categories = categories
        chart_data.add_series('Series 1', iter(frequencies))
        left, top, width, height = Inches(0), Inches(2), Inches(6.5), Inches(5)

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
        if question.stat == 'percent':
            data_labels.number_format = '0%'
        data_labels.position = XL_LABEL_POSITION.OUTSIDE_END
        data_labels.font.size = Pt(12)
        plot = chart.plots[0]
        series = plot.series[0]
        fill = series.format.fill
        fill.solid()
        fill.fore_color.theme_color = MSO_THEME_COLOR.ACCENT_3

    def column_clustered_basic_trended(self, question, slide):
        headers = self.max_years(question.responses)
        series_lists = []
        categories = []
        for header in headers:
            frequencies = []
            series_lists.append(frequencies)
        for response in question.responses:
            categories.append(str(response.response))
            header_idx = 0
            for header in headers:
                frequencies = series_lists[header_idx]
                if response.frequencies.get(header) is not None:
                    freq = response.frequencies.get(header)
                    frequencies.append(freq)
                header_idx += 1
        chart_data = CategoryChartData()
        chart_data.categories = categories
        for idx in range(0, len(headers)):
            chart_data.add_series(str(headers[idx]), iter(series_lists[idx]))

        left, top, width, height = Inches(0), Inches(2), Inches(6.5), Inches(5)

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
        if question.stat == 'percent':
            data_labels.number_format = '0%'
        data_labels.position = XL_LABEL_POSITION.OUTSIDE_END
        data_labels.font.size = Pt(12)
        chart.has_legend = True
        chart.legend.position = XL_LEGEND_POSITION.BOTTOM
        chart.legend.font.size = Pt(12)
        chart.legend.include_in_layout = False

    def bar_clustered_basic(self, question, slide):
        categories = []
        frequencies = []
        for response in question.responses:
            categories.append(str(response.response))
            for year, frequency in response.frequencies.items():
                frequencies.append(frequency)
        chart_data = CategoryChartData()
        chart_data.categories = categories
        chart_data.add_series('Series 1', iter(frequencies))
        left, top, width, height = Inches(6.75), Inches(2), Inches(6.5), Inches(5)

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
        if question.stat == 'percent':
            data_labels.number_format = '0%'
        data_labels.position = XL_LABEL_POSITION.OUTSIDE_END
        data_labels.font.size = Pt(12)
        plot = chart.plots[0]
        series = plot.series[0]
        fill = series.format.fill
        fill.solid()
        fill.fore_color.theme_color = MSO_THEME_COLOR.ACCENT_2

    def bar_clustered_basic_trended(self, question, slide):
        headers = self.max_years(question.responses)
        series_lists = []
        categories = []
        for header in headers:
            frequencies = []
            series_lists.append(frequencies)
        for response in question.responses:
            categories.append(str(response.response))
            header_idx = 0
            for header in headers:
                frequencies = series_lists[header_idx]
                if response.frequencies.get(header) is not None:
                    freq = response.frequencies.get(header)
                    frequencies.append(freq)
                header_idx += 1
        chart_data = CategoryChartData()
        chart_data.categories = categories
        for idx in range(0, len(headers)):
            chart_data.add_series(str(headers[idx]), iter(series_lists[idx]))
        left, top, width, height = Inches(6.75), Inches(2), Inches(6.5), Inches(5)

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
        if question.stat == 'percent':
            data_labels.number_format = '0%'
        data_labels.position = XL_LABEL_POSITION.OUTSIDE_END
        data_labels.font.size = Pt(12)
        chart.has_legend = True
        chart.legend.position = XL_LEGEND_POSITION.BOTTOM
        chart.legend.font.size = Pt(12)
        chart.legend.include_in_layout = False

    def stacked_bar_simple(self, question, slide):
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

        left, top, width, height = Inches(0), Inches(4), Inches(6.5), Inches(3)

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
        if question.stat == 'percent':
            data_labels.number_format = '0%'
        data_labels.font.size = Pt(12)

    def pie_simple(self, question, slide):
        categories = []
        frequencies = []
        for response in question.responses:
            categories.append(str(response.response))
            for year, frequency in response.frequencies.items():
                frequencies.append(frequency)
        chart_data = CategoryChartData()
        chart_data.categories = categories
        chart_data.add_series('Series 1', iter(frequencies))
        left, top, width, height = Inches(6.75), Inches(2), Inches(6.5), Inches(5)

        chart = slide.shapes.add_chart(XL_CHART_TYPE.PIE, left, top, width, height, chart_data).chart
        chart.chart_title.has_text_frame = True
        chart.chart_title.text_frame.text = "Q: " + str(question.prompt)
        chart.chart_title.text_frame.paragraphs[0].font.size = Pt(12)
        chart.has_legend = True
        chart.legend.position = XL_LEGEND_POSITION.BOTTOM
        chart.legend.font.size = Pt(12)
        chart.legend.include_in_layout = False
        chart.plots[0].has_data_labels = True
        data_labels = chart.plots[0].data_labels
        if question.stat == 'percent':
            data_labels.number_format = '0%'
        data_labels.font.size = Pt(12)

    def column_clustered_matrix_by_categories(self, question, slide):
        sub_questions = question.questions
        series_names = []  # scale points
        series_lists = []  # frequencies for each scale point
        categories = []  # responses for matrix
        is_first = True

        for sub_question in sub_questions:
            #categories.append(str(sub_question.prompt) + " (n=" + str(sub_question.n) + ")")
            if is_first:
                for response in sub_question.responses:
                    series_names.append(response.response)
                question.stat = sub_question.stat
            is_first = False

        for name in series_names:
            frequencies = []
            series_lists.append([])

        for sub_question in sub_questions:
            response_idx = 0
            for response in sub_question.responses:
                frequencies = series_lists[response_idx]
                if response.has_frequency is True:
                    for year, frequency in response.frequencies.items():
                        frequencies.append(frequency)
                else:
                    frequencies.append(0)
                response_idx += 1

        chart_data = CategoryChartData()
        chart_data.categories = categories
        for idx in range(0, len(series_names)):
            chart_data.add_series(series_names[idx], iter(series_lists[idx]))

        left, top, width, height = Inches(0), Inches(2), Inches(6.5), Inches(5)
        chart = slide.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, left, top, width, height, chart_data).chart
        chart.chart_title.has_text_frame = True
        chart.chart_title.text_frame.text = "Q: " + str(question.prompt)
        chart.chart_title.text_frame.paragraphs[0].font.size = Pt(12)

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
        if question.stat == 'percent':
            data_labels.number_format = '0%'
        data_labels.position = XL_LABEL_POSITION.OUTSIDE_END
        data_labels.font.size = Pt(12)
        chart.has_legend = True
        chart.legend.position = XL_LEGEND_POSITION.BOTTOM
        chart.legend.font.size = Pt(12)
        chart.legend.include_in_layout = False

    def column_clustered_matrix_by_scale_points(self, question, slide):
        sub_questions = question.questions
        series_names = []  # responses
        series_lists = []  # frequencies for each scale point
        categories = []  # scale points
        is_first = True

        for sub_question in sub_questions:
            #series_names.append(str(sub_question.prompt) + " (n=" + str(sub_question.n) + ")")
            if is_first:
                for response in sub_question.responses:
                    categories.append(response.response)
                question.stat = sub_question.stat
            is_first = False

        for name in series_names:
            frequencies = []
            series_lists.append([])
        response_idx = 0
        for sub_question in sub_questions:
            frequencies = series_lists[response_idx]
            for response in sub_question.responses:
                if response.has_frequency is True:
                    for year, frequency in response.frequencies.items():
                        frequencies.append(frequency)
                else:
                    frequencies.append(0)
            response_idx += 1

        chart_data = CategoryChartData()
        chart_data.categories = categories
        for idx in range(0, len(series_names)):
            chart_data.add_series(series_names[idx], iter(series_lists[idx]))

        left, top, width, height = Inches(6.75), Inches(2), Inches(6.5), Inches(5)
        chart = slide.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, left, top, width, height, chart_data).chart
        chart.chart_title.has_text_frame = True
        chart.chart_title.text_frame.text = "Q: " + str(question.prompt)
        chart.chart_title.text_frame.paragraphs[0].font.size = Pt(12)

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
        if question.stat == 'percent':
            data_labels.number_format = '0%'
        data_labels.position = XL_LABEL_POSITION.OUTSIDE_END
        data_labels.font.size = Pt(12)
        chart.has_legend = True
        chart.legend.position = XL_LEGEND_POSITION.BOTTOM
        chart.legend.font.size = Pt(12)
        chart.legend.include_in_layout = False

    def bar_clustered_matrix_by_categories(self, question, slide):
        sub_questions = question.questions
        series_names = []  # scale points
        series_lists = []  # frequencies for each scale point
        categories = []  # responses for matrix
        is_first = True

        for sub_question in sub_questions:
            #categories.append(str(sub_question.prompt) + " (n=" + str(sub_question.n) + ")")
            if is_first:
                for response in sub_question.responses:
                    series_names.append(response.response)
                question.stat = sub_question.stat
            is_first = False

        for name in series_names:
            frequencies = []
            series_lists.append([])

        for sub_question in sub_questions:
            response_idx = 0
            for response in sub_question.responses:
                frequencies = series_lists[response_idx]
                if response.has_frequency is True:
                    for year, frequency in response.frequencies.items():
                        frequencies.append(frequency)
                else:
                    frequencies.append(0)
                response_idx += 1

        chart_data = CategoryChartData()
        chart_data.categories = categories
        for idx in range(0, len(series_names)):
            chart_data.add_series(series_names[idx], iter(series_lists[idx]))
        left, top, width, height = Inches(0), Inches(2), Inches(6.5), Inches(5)

        chart = slide.shapes.add_chart(XL_CHART_TYPE.BAR_CLUSTERED, left, top, width, height, chart_data).chart
        chart.chart_title.has_text_frame = True
        chart.chart_title.text_frame.text = "Q: " + str(question.prompt)
        chart.chart_title.text_frame.paragraphs[0].font.size = Pt(12)

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
        if question.stat == 'percent':
            data_labels.number_format = '0%'
        data_labels.position = XL_LABEL_POSITION.OUTSIDE_END
        data_labels.font.size = Pt(12)
        chart.has_legend = True
        chart.legend.position = XL_LEGEND_POSITION.BOTTOM
        chart.legend.font.size = Pt(12)
        chart.legend.include_in_layout = False

    def bar_clustered_matrix_by_scale_points(self, question, slide):
        sub_questions = question.questions
        series_names = []  # responses
        series_lists = []  # frequencies for each scale point
        categories = []  # scale points
        is_first = True

        for sub_question in sub_questions:
            #series_names.append(str(sub_question.prompt) + " (n=" + str(sub_question.n) + ")")
            if is_first:
                for response in sub_question.responses:
                    categories.append(response.response)
                question.stat = sub_question.stat
            is_first = False

        for name in series_names:
            frequencies = []
            series_lists.append([])
        response_idx = 0
        for sub_question in sub_questions:
            frequencies = series_lists[response_idx]
            for response in sub_question.responses:
                if response.has_frequency is True:
                    for year, frequency in response.frequencies.items():
                        frequencies.append(frequency)
                else:
                    frequencies.append(0)
            response_idx += 1

        chart_data = CategoryChartData()
        chart_data.categories = categories
        for idx in range(0, len(series_names)):
            chart_data.add_series(series_names[idx], iter(series_lists[idx]))
        left, top, width, height = Inches(6.75), Inches(2), Inches(6.5), Inches(5)

        chart = slide.shapes.add_chart(XL_CHART_TYPE.BAR_CLUSTERED, left, top, width, height, chart_data).chart
        chart.chart_title.has_text_frame = True
        chart.chart_title.text_frame.text = "Q: " + str(question.prompt)
        chart.chart_title.text_frame.paragraphs[0].font.size = Pt(12)

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
        if question.stat == 'percent':
            data_labels.number_format = '0%'
        data_labels.position = XL_LABEL_POSITION.OUTSIDE_END
        data_labels.font.size = Pt(12)
        chart.has_legend = True
        chart.legend.position = XL_LEGEND_POSITION.BOTTOM
        chart.legend.font.size = Pt(12)
        chart.legend.include_in_layout = False

    def stacked_bar_matrix(self, question, slide, side):
        sub_questions = question.questions
        series_names = []  # scale points
        series_lists = []  # frequencies for each scale point
        categories = []  # responses for matrix
        is_first = True

        for sub_question in sub_questions:
            #categories.append(str(sub_question.prompt) + " (n=" + str(sub_question.n) + ")")
            if is_first:
                for response in sub_question.responses:
                    series_names.append(response.response)
                question.stat = sub_question.stat
            is_first = False

        for name in series_names:
            frequencies = []
            series_lists.append([])

        for sub_question in sub_questions:
            response_idx = 0
            for response in sub_question.responses:
                frequencies = series_lists[response_idx]
                if response.has_frequency is True:
                    for year, frequency in response.frequencies.items():
                        frequencies.append(frequency)
                else:
                    frequencies.append(0)
                response_idx += 1

        chart_data = CategoryChartData()
        chart_data.categories = categories
        for idx in range(0, len(series_names)):
            chart_data.add_series(series_names[idx], iter(series_lists[idx]))

        if side == 'top':
            left, top, width, height = Inches(1.5), Inches(1), Inches(10), Inches(3)
        if side == 'bottom':
            left, top, width, height = Inches(1.5), Inches(4), Inches(10), Inches(3)

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
        if question.stat == 'percent':
            data_labels.number_format = '0%'
        data_labels.font.size = Pt(12)
        plot = chart.plots[0]
        num_series = len(plot.series)


        if num_series <= 6:
            colors = []
            if num_series >= 1:
                colors.append(1)
            if num_series >= 2:
                colors.append(6)
            if num_series >= 3:
                colors.insert(1, 4)
            if num_series >= 4:
                colors.insert(1, 2)
            if num_series >= 5:
                colors.insert(3, 5)
            if num_series == 6:
                colors.insert(2, 3)

            if side == 'bottom':
                colors.reverse()



            color_index = 0
            for series in plot.series:
                fill = series.format.fill
                fill.solid()
                if colors[color_index] == 1:
                    fill.fore_color.theme_color = MSO_THEME_COLOR.ACCENT_1
                elif colors[color_index] == 2:
                    fill.fore_color.theme_color = MSO_THEME_COLOR.ACCENT_2
                elif colors[color_index] == 3:
                    fill.fore_color.theme_color = MSO_THEME_COLOR.ACCENT_3
                elif colors[color_index] == 4:
                    fill.fore_color.theme_color = MSO_THEME_COLOR.ACCENT_4
                elif colors[color_index] == 5:
                    fill.fore_color.theme_color = MSO_THEME_COLOR.ACCENT_5
                elif colors[color_index] == 6:
                    fill.fore_color.theme_color = MSO_THEME_COLOR.ACCENT_6
                color_index += 1
        else:
            print("CANNOT REVERSE GRADIENT FOR STACKED BARS WITH MORE THAN 6 SERIES")




    def column_clustered_allocation(self, question, slide):
        sub_questions = question.questions
        categories = []
        frequencies = []

        for sub_question in sub_questions:
            for response in sub_question.responses:
                #categories.append(response.response + " (n=" + str(sub_question.n) + ")")
                for year, frequency in response.frequencies.items():
                    frequencies.append(frequency)

        chart_data = CategoryChartData()
        chart_data.categories = categories
        chart_data.add_series('Series 1', iter(frequencies))
        left, top, width, height = Inches(0), Inches(2), Inches(6.5), Inches(5)

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
        data_labels.position = XL_LABEL_POSITION.OUTSIDE_END
        data_labels.font.size = Pt(12)
        plot = chart.plots[0]
        series = plot.series[0]
        fill = series.format.fill
        fill.solid()
        fill.fore_color.theme_color = MSO_THEME_COLOR.ACCENT_1

    def bar_clustered_allocation(self, question, slide):
        sub_questions = question.questions
        categories = []
        frequencies = []

        for sub_question in sub_questions:
            for response in sub_question.responses:
                #categories.append(response.response + " (n=" + str(sub_question.n) + ")")
                for year, frequency in response.frequencies.items():
                    frequencies.append(frequency)

        chart_data = CategoryChartData()
        chart_data.categories = categories
        chart_data.add_series('Series 1', iter(frequencies))
        left, top, width, height = Inches(6.75), Inches(2), Inches(6.5), Inches(5)

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
        # data_labels.number_format = '0%'
        data_labels.position = XL_LABEL_POSITION.OUTSIDE_END
        data_labels.font.size = Pt(12)
        plot = chart.plots[0]
        series = plot.series[0]
        fill = series.format.fill
        fill.solid()
        fill.fore_color.theme_color = MSO_THEME_COLOR.ACCENT_1

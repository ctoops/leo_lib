from openpyxl import load_workbook
from openpyxl.chart import PieChart, Reference, BarChart
from openpyxl.chart.label import DataLabelList
from openpyxl.chart.marker import DataPoint
from openpyxl.chart.text import RichText
from openpyxl.drawing.colors import ColorChoice
from openpyxl.drawing.text import CharacterProperties, Paragraph, ParagraphProperties


class Chart:

    def __init__(self, file=None, data_only: bool = False, read_only: bool = False, keep_vba: bool = False):
        self.workbook = load_workbook(filename=file, data_only=data_only, read_only=read_only, keep_vba=keep_vba)
        self.cp = CharacterProperties(solidFill=ColorChoice(prstClr='white'), sz=1400)
        self.chart_width = 14.77
        self.chart_height = 7.65
        self.chart_row = 6

    def pie_chart(self, sheet_name, label_range, data_range, position, title, color: list = None,
                  from_rows: bool = False, titles_from_data: bool = False, label_options: DataLabelList = None,
                  chart_size: tuple = None):
        """基础饼图

        :param sheet_name: sheet名
        :type sheet_name str
        :param label_range: 饼图label来源，格式为：sheetname!A1:E4
        :type label_range str
        :param data_range: 饼图数据来源，格式为： sheetname!A1:E4
        :type data_range str
        :param position: 饼图的位置，格式为：A1
        :param title: 饼图的标题
        :param color: 饼图每个区域的颜色，
        :type color list 每个区域的16进制颜色列表，按label来源排序
        :param from_rows: 是否从行读取label， True: 行读取label， False 列读取label
        :param titles_from_data: 默认为False
        :type titles_from_data bool
        :param label_options: 饼图中显示的文本的属性
        :type label_options DataLabelList
        :param chart_size: 饼图尺寸，(width, height)
        :type chart_size tuple
        :return:
        """
        ws = self.workbook[sheet_name]
        pie = PieChart()
        label_source = Reference(range_string=label_range)
        data_source = Reference(range_string=data_range)
        pie.add_data(data_source, from_rows=from_rows, titles_from_data=titles_from_data)
        pie.set_categories(label_source)
        pie.title = title
        s = pie.series[0]
        if color:
            for index, item in enumerate(color):
                dp = DataPoint(idx=index)
                dp.graphicalProperties.solidFill = item
                s.dPt.append(dp)
        if label_options:
            pie.dataLabels = label_options
        else:
            pie.dataLabels = DataLabelList(dLblPos='bestFit', txPr=RichText(
                p=[Paragraph(pPr=ParagraphProperties(defRPr=self.cp), endParaRPr=self.cp)]), showPercent=True)
        if chart_size:
            pie.width = chart_size[0]
            pie.height = chart_size[1]
        ws.add_chart(pie, position)

    def bar_chart(self, sheet_name, label_range, data_range, position, title, x_title: str = None,
                  y_title: str = None, bar_type: str = 'col', color: list = None,
                  from_rows: bool = False, titles_from_data: bool = False,
                  label_options: DataLabelList = None, chart_size: tuple = None, legend: str = None):
        """基础柱状图/ 直方图，一个柱体只有一个数据来源

        :param sheet_name:
        :param label_range: 柱状图label来源，格式为：sheetname!A1:E4
        :param data_range: 柱状图数据来源，格式为： sheetname!A1:E4
        :param position: 柱状图的位置，格式为：A1
        :param title: 柱状图的标题
        :param x_title: 柱状图X轴标题
        :param y_title: 柱状图Y轴标题
        :param legend: 柱体之间的距离
        :param bar_type:
        :param color: 每个区域的16进制颜色列表，按label来源排序
        :param from_rows: label和data的来源是否从行读取
        :param titles_from_data: 默认为False
        :param label_options: 柱状图中label的属性
        :param chart_size: 柱状图的尺寸 (width, height)
        :return:
        """
        ws = self.workbook[sheet_name]
        bar = BarChart()
        label_source = Reference(range_string=label_range)
        data_source = Reference(range_string=data_range)
        bar.add_data(data_source, from_rows=from_rows, titles_from_data=titles_from_data)
        bar.set_categories(label_source)
        bar.title = title
        bar.type = bar_type
        if x_title:
            bar.x_axis.title = x_title
        if y_title:
            bar.y_axis.title = y_title
        b_s = bar.series[0]
        if color:
            for index, item in enumerate(color):
                dp = DataPoint(idx=index)
                dp.graphicalProperties.solidFill = item
                b_s.dPt.append(dp)
        if label_options:
            bar.dataLabels = label_options
        else:
            bar.dataLabels = DataLabelList(dLblPos='ctr', txPr=RichText(
                p=[Paragraph(pPr=ParagraphProperties(defRPr=self.cp), endParaRPr=self.cp)]), showVal=True)
        bar.legend = legend
        if chart_size:
            bar.width = chart_size[0]
            bar.height = chart_size[1]
        ws.add_chart(bar, position)

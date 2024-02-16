import openpyxl
from openpyxl.drawing.image import Image
from openpyxl.drawing.spreadsheet_drawing import AnchorMarker, TwoCellAnchor
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.utils import get_column_letter
from openpyxl.utils.dataframe import dataframe_to_rows
from pandas import DataFrame

fill_head = PatternFill('solid', fgColor='8db4e2')  # 设置填充颜色为 灰色
font_head = Font(u'Calibri', size=14, color='000000', bold=True,
                 italic=False, strike=False)  # 设置字体样式
font_body = Font(u'Calibri', size=12, color='000000', bold=False,
                 italic=False, strike=False)  # 设置字体样式
align = Alignment(horizontal='center', vertical='center', wrapText=True)


class MLAlertTable:

    def __init__(self, test_info: DataFrame):
        self.workbook = openpyxl.Workbook()
        self.worksheet = self.workbook.active
        self.test_info = test_info
        self.headers = []
        self.fill_data_frame()
        self.set_header()
        self.set_body()

    def fill_data_frame(self):
        for row in dataframe_to_rows(self.test_info, index=False, header=True):
            self.worksheet.append(row)

    def set_header(self):
        self.worksheet.row_dimensions[1].height = 40

        for i in range(self.worksheet.max_column):
            col_letter = get_column_letter(i + 1)
            ccw = self.worksheet.column_dimensions[col_letter].width
            head = self.worksheet[f'{col_letter}1']
            head.fill = fill_head
            head.font = font_head
            head.alignment = align

            if ccw < 27 and head.value.endswith('_pic'):
                self.worksheet.column_dimensions[col_letter].width = 27
            elif ccw < 20:
                self.worksheet.column_dimensions[col_letter].width = 20
            self.headers.append(head.value)

    def set_body(self):
        for j in range(2, self.worksheet.max_row + 1):
            self.worksheet.row_dimensions[j].height = 85
            for i in range(1, self.worksheet.max_column + 1):
                head = self.headers[i - 1]
                cell = self.worksheet.cell(j, i)
                cell.alignment = align
                cell.font = font_body
                if head.endswith('_pic') and cell.value != 'NA':
                    self.insert_image_to_cell(j - 1, i - 1, cell.value)
                    cell.value = ''

    def get_cell_width_height(self, row, col) -> tuple:
        col_letter = get_column_letter(col + 1)
        ccw = self.worksheet.column_dimensions[col_letter].width
        crh = self.worksheet.row_dimensions[row + 1].height
        w = (ccw * 72) // 9
        h = (crh * 18) // 13.5
        r = w / h
        return w, h, r

    def insert_image_to_cell(self, row, col, image_url):
        img = Image(image_url)
        _from = AnchorMarker(col, 100000, row, 100000)
        _to = AnchorMarker(col + 1, -100000, row + 1, -100000)
        img.anchor = TwoCellAnchor('twoCell', _from, _to)
        self.worksheet.add_image(img)


if __name__ == '__main__':
    pass

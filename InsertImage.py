from openpyxl.drawing.image import Image
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.drawing.spreadsheet_drawing import AnchorMarker, TwoCellAnchor
from openpyxl.utils.units import pixels_to_EMU
from openpyxl.utils import get_column_letter
from openpyxl.utils.dataframe import dataframe_to_rows
import openpyxl
from pathlib import Path
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
            print(col_letter)
            ccw = self.worksheet.column_dimensions[col_letter].width
            head = self.worksheet[f'{col_letter}1']
            head.fill = fill_head
            head.font = font_head
            head.alignment = align

            if ccw < 35 and head.value.endswith('JPG'):
                self.worksheet.column_dimensions[col_letter].width = 35
            elif ccw < 20:
                self.worksheet.column_dimensions[col_letter].width = 20

    def set_body(self):
        for j in range(1, self.worksheet.max_row + 1):
            self.worksheet.row_dimensions[j].height = 85
            for i in range(1, self.worksheet.max_column + 1):
                cell = self.worksheet.cell(j, i)
                cell.alignment = align
                cell.font = font_body

    def insert_image_to_cell(self, row, col, image_url):
        img = Image(image_url)
        w, h = img.width, img.height
        ratio = w / h
        print(w, h, ratio)
        x, y, ww, hh = (325.0, 0.0, 240, 113.0)
        print(x, y, ww, hh)
        hhh = hh - 6
        www = int(hhh * ratio)
        xx = x + (ww - www) // 2
        yy = y + 3
        _from = AnchorMarker(col, 50000, row, 50000)
        _to = AnchorMarker(col + 1, -50000, row + 1, -50000)
        img.size = (www, hhh)
        img.anchor = TwoCellAnchor('twoCell', _from, _to)
        self.worksheet.add_image(img)


if __name__ == '__main__':
    add_image = MLAlertTable()
    print(add_image.insert_image_to_cell(1, 3,
                                         'D:/Users/xiaoze_wang/Desktop/ExportBlob_v3.0/ParseFile/DataZip/Data/C441NW3L6W_FAIL_20240207023114/D83.C441NW3L6W.BGS.20240207.023457.BGI.JPG',
                                         ))
    add_image.workbook.save(r'D:\Users\Xiaoze_Wang\Desktop\BGS ML Alert SN re-judge tracker list2.xlsx')

from openpyxl.drawing.image import Image
from openpyxl.drawing.xdr import XDRPoint2D, XDRPositiveSize2D
from openpyxl.drawing.spreadsheet_drawing import AbsoluteAnchor
from openpyxl.utils.units import pixels_to_EMU
from openpyxl.utils import get_column_letter
import openpyxl


class AddImage():

    def __init__(self):
        self.workbook = openpyxl.Workbook()
        self.worksheet = self.workbook.active

    def get_absolute(self, row, col):
        """
        获取单元格的右下方绝对位置（单位：像素），及单元格的宽高
        """
        x = 0
        y = 0
        # get_column_letter(int)把整数转换为Excel中的列索引
        col_letter = get_column_letter(col)
        # 获取每列的列宽
        width = self.worksheet.column_dimensions[col_letter].width
        # 计算第一列到目标列的总宽
        for i in range(col):
            col_letter = get_column_letter(i + 1)
            fcw = self.worksheet.column_dimensions[col_letter].width
            x += fcw
        # 如果Excel中高为默认值时，openpyxl却没有值为NoneValue，这一点我很奇怪。
        if not self.worksheet.row_dimensions[col].height:
            self.worksheet.row_dimensions[col].height = 13.5
            height = 13.5  # Excel默认列宽为13.5
        else:
            height = self.worksheet.row_dimensions[col].height
        # 计算第一行到目标行的总高
        for j in range(row):
            if not self.worksheet.row_dimensions[j + 1].height:
                self.worksheet.row_dimensions[j + 1].height = 13.5
                fch = 13.5
            else:
                fch = self.worksheet.row_dimensions[j + 1].height
            y += fch
            # 把高单位转换为像素
        height = (height * 18) // 13.5  # 一个单元格高为13.5，像素为18
        # 把宽单位转换为像素
        width = (width * 72) // 9  # 一个单元格为宽为9，像素为72
        x = (x * 72) // 9
        y = (y * 18) // 13.5
        return x, y, width, height

    def insert_image(self, row, col, end_row, end_col, image_url, image_size=None):

        img = Image(image_url)
        if image_size:
            img.width, img.height = image_size
        w, h = img.width, img.height
        x1, y1, w1, h1 = self.get_absolute(row, col)
        x2, y2, w2, h2 = self.get_absolute(end_row, end_col)
        x = (x2 + x1 - w1 - h) // 2
        y = (y2 + y1 - h1 - w) // 2
        p2e = pixels_to_EMU  # openpylx自带的像素转EMU
        pos = XDRPoint2D(p2e(x), p2e(y))  # 设置绝对位置
        size = XDRPositiveSize2D(p2e(w), p2e(h))  # 图片大小
        img.anchor = AbsoluteAnchor(pos=pos, ext=size)
        self.worksheet.add_image(img)


if __name__ == '__main__':
    pass

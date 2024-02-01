# @File: Create_Template.py
# @Time: 2024/1/28 下午 10:32  
# @Author: Nan1_Chen
# @Mail: Nan1_Chen@pegatroncorp.com

import re
import openpyxl
import pandas as pd
from PIL import Image as PILImage
from openpyxl.drawing.image import Image
from openpyxl.styles import Font
from openpyxl.styles import PatternFill, Alignment
from pathlib import Path
from LoadFilePath import data_pardir,export_dir,alert_fail_units_dir,alert_file_path
from ExtractZIP import file_name


excel_path = export_dir / "XGS ML Alert SN.xlsx"


def create_template():
    wb = openpyxl.Workbook()
    ws = wb.active

    # 格式設置
    ws.row_dimensions[1].height = 40
    fille = PatternFill('solid', fgColor='adaaa6')  # 设置填充颜色为 橙色
    fille_NG = PatternFill('solid', fgColor='ff99cc')  # 设置填充颜色为 紅色
    fille_OK = PatternFill('solid', fgColor='AACF91')  # 设置填充颜色为 綠色
    font = Font(u'Calibri', size=12, color='ffffff', bold=True, italic=False, strike=False)  # 设置字体样式
    ws['a1'].fill = fille  # 应用填充样式在A1单元格
    ws['b1'].fill = fille  # 应用填充样式在B1单元格
    ws['c1'].fill = fille  # 应用填充样式在C1单元格
    ws['d1'].fill = fille  # 应用填充样式在D1单元格
    ws['e1'].fill = fille  # 应用填充样式在E1单元格
    ws['f1'].fill = fille  # 应用填充样式在F1单元格
    ws['g1'].fill = fille  # 应用填充样式在G1单元格
    ws['h1'].fill = fille  # 应用填充样式在H1单元格
    ws['i1'].fill = fille  # 应用填充样式在I1单元格
    ws['j1'].fill = fille  # 应用填充样式在J1单元格
    ws['k1'].fill = fille  # 应用填充样式在J1单元格
    ws['a1'].font = font  # 应用填充样式在A1单元格
    ws['b1'].font = font  # 应用填充样式在B1单元格
    ws['c1'].font = font  # 应用填充样式在C1单元格
    ws['d1'].font = font  # 应用填充样式在D1单元格
    ws['e1'].font = font  # 应用填充样式在E1单元格
    ws['f1'].font = font  # 应用填充样式在F1单元格
    ws['g1'].font = font  # 应用填充样式在G1单元格
    ws['h1'].font = font  # 应用填充样式在H1单元格
    ws['i1'].font = font  # 应用填充样式在I1单元格
    ws['j1'].font = font  # 应用填充样式在J1单元格
    ws['k1'].font = font  # 应用填充样式在J1单元格
    # 向excel中写入表头
    ws['a1'] = 'SN'
    ws['b1'] = 'Station ID'
    ws['c1'] = 'Test Time'
    ws['d1'] = 'BGI Image'
    ws['e1'] = 'ICE Image'
    ws['f1'] = 'SOB Image'
    ws['g1'] = 'BGI Issue'
    ws['h1'] = 'ICE Issue'
    ws['i1'] = 'SOB Issue'
    ws['j1'] = 'Root Cause'
    ws['k1'] = 'CA'
    ws.column_dimensions['A'].width = 15
    ws.column_dimensions['B'].width = 25
    ws.column_dimensions['C'].width = 15
    ws.column_dimensions['D'].width = 25
    ws.column_dimensions['E'].width = 25
    ws.column_dimensions['F'].width = 25
    ws.column_dimensions['G'].width = 15
    ws.column_dimensions['H'].width = 15
    ws.column_dimensions['I'].width = 15
    ws.column_dimensions['J'].width = 30
    ws.column_dimensions['K'].width = 30
    # 加載albrt_file
    df = pd.read_csv(alert_file_path, usecols=["serial_number", "station_id", "uut_start", "model_sob_decision"
        , "model_ice_decision", "model_bgi_decision"])

    # 插入信息：
    x = 2  # 在第一行开始写
    y = 1  # 在第一列开始写
    kk = 1

    for unit_folder_path in data_pardir.iterdir():
        print(unit_folder_path)
        Units_Folder_Name = unit_folder_path.name  # 返回文件名
        # print(Units_Folder_Name)
        SN = Units_Folder_Name.split('_')
        # print(SN)
        print(f"*****第 {kk} 個機臺開始插入信息************************************************************")
        # 1.插入第一欄SN信息：
        print(f"-----1.正在插入A欄 Serial Number 信息: {SN[0]}---------")
        ws.cell(row=x, column=y).value = str(SN[0])
        ii = 0
        for item in df["serial_number"]:
            if item == SN[0]:
                # 2.插入第二欄Station ID信息：
                print(f"-----2.正在插入B欄 Station ID 信息: {str(df['station_id'][ii])}----------------")
                ws.cell(row=x, column=y + 1).value = str(df["station_id"][ii])  # Station ID
                # 3.插入第三欄Test Time信息：
                print(f"-----3.正在插入C欄 Test Time 信息: {str(df['uut_start'][ii])} -------------")
                ws.cell(row=x, column=y + 2).value = str(df["uut_start"][ii])  # Time
            ii += 1

        # 4.插入第四欄BGI Image信息：
        print('-----4.正在插入D欄 BGI圖片--------------------')

        for path in file_name(unit_folder_path,r'.*(SOBBK|ICEBK|BGI)\.JPG'):
            image = str(path)  # 返回文件名
            image_type = re.findall(r'.*(SOBBK|ICEBK|BGI)\.JPG',path.name)[0]
            im = PILImage.open(image)
            w, h = im.size
            im.thumbnail((w // 3, h // 3))
            im.save(image)
            img_column = {'SOBBK':'F','ICEBK':'E','BGI':'D'}.get(image_type,'D')
            # 图片路径
            img_file_path = image
            # 获取图片
            img = Image(image)
            # 设置图片的大小
            img.width, img.height = (110, 110)
            # 设置表格的宽20和高85
            ws.column_dimensions[img_column].width = 20
            ws.row_dimensions[x].height = 85
            # 图片插入名称对应单元格
            ws.add_image(img, anchor=img_column + str(x))

        # 7.插入第七欄BGI Issue信息：
        print('-----7.正在插入H欄 BGI Ml infor--------------------')
        for text_path in file_name(unit_folder_path,r'.*(SOBBK|ICEBK|BGI)\.txt'):
            text = str(text_path)  # 返回文件名
            text_type = re.findall(r'.*(SOBBK|ICEBK|BGI)\.txt',text_path.name)[0]
            y_offset = {'SOBBK': 8, 'ICEBK': 7, 'BGI': 6}.get(text_type, '6')
            with open(text, "r", encoding='utf-8') as f:
                content = f.read()
                result1 = re.findall('.*,"issues":\[(.*)],.*', content)
                result1 = " ".join(result1)

                if not result1.strip():
                    # print("該機臺bgi無異常")
                    ws.cell(row=x, column=y + y_offset).value = "no issue"
                    ws.cell(row=x, column=y + y_offset).fill = fille_OK
                else:
                    result1 = eval(result1)
                    # print(result1)
                    ws.cell(row=x, column=y + y_offset).value = str(result1)
                    ws.cell(row=x, column=y + y_offset).fill = fille_NG

        print(f"*****第 {kk} 個機臺信息插入完成************************************************************")
        print("                                                     ")
        x += 1
        kk += 1

    max_rows = ws.max_row  # 获取最大行
    max_columns = ws.max_column  # 获取最大列
    align = Alignment(horizontal='center', vertical='center')
    # openpyxl的下标从1开始
    for iiii in range(1, max_rows + 1):
        for jjjj in range(1, max_columns + 1):
            ws.cell(iiii, jjjj).alignment = align
    wb.save(excel_path)
    wb.close()


if __name__ == "__main__":

    create_template()

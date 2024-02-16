# @File: main.py
# @Time: 2024/1/28 上午 10:25
# @Author: Nan1_Chen
# @Mail: Nan1_Chen@pegatroncorp.com
import datetime
import Create_Template
import ExtractZIP
import LoadFilePath
import MarkPoint

current_start_time = datetime.datetime.now()


def main():
    print(f"開始時間：{current_start_time}")
    LoadFilePath.init_folders()
    LoadFilePath.select_file()
    ExtractZIP.decompression_ml_image_txt()
    MarkPoint.mark_bgi_ml_image()
    Create_Template.create_template()
    current_end_time = datetime.datetime.now()
    print(f"結束時間：{current_end_time}")
    a = round(((current_end_time - current_start_time).seconds / 60), 2)
    print(f"總共運行時間：{a} 分鐘")


main()

if "__name__" == "__main__":
    main()

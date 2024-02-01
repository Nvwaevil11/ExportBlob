# @File: MarkPiont.py
# @Time: 2024/1/28 下午 02:55  
# @Author: Nan1_Chen
# @Mail: Nan1_Chen@pegatroncorp.com

# @Software: PyCharm
# --- --- --- --- --- --- --- --- ---


import os
import re
from glob import glob
from LoadFilePath import alert_fail_units_dir,data_dir,alert_fail_units_pardir,data_pardir,data_zippardir
from ExtractZIP import file_name
import cv2


def MarkSOBimage():
    paths = glob(data_pardir + r'//*//*//*//*SOBBK.txt')
    kk = 1
    for path in paths:
        with open(path, "r", encoding='utf-8') as f:
            content = f.read()
            # ,"issues":["foreign_material"],
            # ,"issues":["leakage"],
            result = re.findall('.*,"issues":\[(.*)],.*', content)
            result = " ".join(result)
            if not result.strip():
                print("該機臺sob無異常")
            else:
                print(eval(result))
                with open(path, "r", encoding='utf-8') as f:
                    content = f.read()
                    result1 = re.findall('.*"bbox":\{(.*)},.*', content)
                    result1 = " ".join(result1)
                    # print(result)
                    x1 = int(result1.split(':')[1].split(',')[0])
                    y1 = int(result1.split(':')[2].split(',')[0])
                    x2 = int(result1.split(':')[3].split(',')[0])
                    y2 = int(result1.split(':')[4].split(',')[0])
                    # print(x1, y1, x2, y2)
                    x3 = int(x2) + int(y1)
                    y3 = int(x1) + int(y2)
                    print(f"異常位置的坐標：{x2, y2},{x3, y3}")
                    SOB_PATH = path.split('.txt')[0] + '.JPG'
                    SOB_StationName = SOB_PATH[-40:-30]
                    print(SOB_StationName)
                    img = cv2.imread(SOB_PATH)
                    # time.sleep(3)
                    cv2.rectangle(img, (int(x2), int(y2)), (x3, y3), (0, 0, 255), 2)
                    font = cv2.FONT_HERSHEY_SIMPLEX
                    cv2.putText(img, eval(result), (int(x2) - 25, int(y2) - 25), font, 1.5, (0, 0, 255), 3)
                    cv2.putText(img, SOB_StationName, (200, 100), font, 3, (255, 0, 0), 5)
                    # cv2.namedWindow('SOB', 0)
                    # cv2.imshow("SOB", img)
                    # cv2.waitKey(0)
                    # cv2.destroyAllWindows()
                    cv2.imwrite(SOB_PATH, img)
        print(f"*****第 {kk} 個機臺信息檢查完成************************************************************")
        print("                                                     ")
        kk += 1


# MarkSOBimage()


def MarkICEimage():
    paths = data_pardir.glob(r'//*//*//*//*ICEBK.txt')
    kk = 1
    for path in paths:
        with open(path, "r", encoding='utf-8') as f:
            content = f.read()
            # ,"issues":["foreign_material"],
            # ,"issues":["leakage"],
            result = re.findall('.*,"issues":\[(.*)],.*', content)
            result = " ".join(result)
            if not result.strip():
                print("該機臺ice無異常")
            else:
                print(eval(result))
                with open(path, "r", encoding='utf-8') as f:
                    content = f.read()
                    result1 = re.findall('.*"bbox":\{(.*)},.*', content)
                    result1 = " ".join(result1)
                    # print(result)
                    x1 = int(result1.split(':')[1].split(',')[0])
                    y1 = int(result1.split(':')[2].split(',')[0])
                    x2 = int(result1.split(':')[3].split(',')[0])
                    y2 = int(result1.split(':')[4].split(',')[0])
                    # print(x1, y1, x2, y2)
                    x3 = int(x2) + int(y1)
                    y3 = int(x1) + int(y2)
                    print(f"異常位置的坐標：{x2, y2},{x3, y3}")
                    SOB_PATH = path.split('.txt')[0] + '.JPG'
                    SOB_StationName = SOB_PATH[-40:-30]
                    img = cv2.imread(SOB_PATH)
                    cv2.rectangle(img, (int(x2), int(y2)), (x3, y3), (0, 0, 255), 2)
                    font = cv2.FONT_HERSHEY_SIMPLEX
                    cv2.putText(img, eval(result), (int(x2) - 25, int(y2) - 25), font, 1.5, (0, 0, 255), 3)
                    cv2.putText(img, SOB_StationName, (200, 100), font, 3, (255, 0, 0), 5)
                    # cv2.namedWindow('SOB', 0)
                    # cv2.imshow("SOB", img)
                    # cv2.waitKey(0)
                    # cv2.destroyAllWindows()
                    cv2.imwrite(SOB_PATH, img)
        print(f"*****第 {kk} 個機臺信息檢查完成************************************************************")
        print("                                                     ")
        kk += 1


# MarkICEimage()


def MarkBGIimage():
    kk = 1
    for path in file_name(data_pardir,r'.*(SOBBK|ICEBK|BGI)\.txt'):
        with open(path, "r", encoding='utf-8') as f:
            content = f.read()
            # ,"issues":["foreign_material"],
            # ,"issues":["leakage"],
            result = re.findall('.*,"issues":\[(.*)],.*', content)
            result = " ".join(result)
            if not result.strip():
                print("該機臺bgi無異常")
            else:
                print(eval(result))
                with open(path, "r", encoding='utf-8') as f:
                    content = f.read()
                    result1 = re.findall('.*"bbox":\{(.*)},.*', content)
                    result1 = " ".join(result1)
                    # print(result)
                    x1 = int(result1.split(':')[1].split(',')[0])
                    y1 = int(result1.split(':')[2].split(',')[0])
                    x2 = int(result1.split(':')[3].split(',')[0])
                    y2 = int(result1.split(':')[4].split(',')[0])
                    # print(x1, y1, x2, y2)
                    x3 = int(x2) + int(y1)
                    y3 = int(x1) + int(y2)
                    print(f"異常位置的坐標：{x2, y2},{x3, y3}")
                    SOB_PATH =path.parent / ( path.stem + '.JPG')
                    # print(ppp)
                    SOB_StationName = SOB_PATH.stem[-38:-28]
                    img = cv2.imread(str(SOB_PATH))
                    cv2.rectangle(img, (int(x2), int(y2)), (x3, y3), (0, 0, 255), 2)
                    font = cv2.FONT_HERSHEY_SIMPLEX
                    cv2.putText(img, eval(result), (int(x2) - 25, int(y2) - 25), font, 1.5, (0, 0, 255), 3)
                    cv2.putText(img, SOB_StationName, (200, 100), font, 3, (255, 0, 0), 5)
                    # cv2.namedWindow('SOB', 0)
                    # cv2.imshow("SOB", img)
                    # cv2.waitKey(0)
                    # cv2.destroyAllWindows()
                    cv2.imwrite(str(SOB_PATH), img)
        print(f"*****第 {kk} 個機臺信息檢查完成************************************************************")
        print("                                                     ")
        kk += 1

# MarkBGIimage()
#
# if "__name__" == "__main__":
#     MarkSOBimage()
#     MarkICEimage()
#     MarkBGIimage()

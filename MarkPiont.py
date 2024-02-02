# @File: MarkPiont.py
# @Time: 2024/1/28 下午 02:55  
# @Author: Nan1_Chen
# @Mail: Nan1_Chen@pegatroncorp.com

# @Software: PyCharm
# --- --- --- --- --- --- --- --- ---

from math import log2
import re
from LoadFilePath import data_pardir
from ExtractZIP import file_name
import cv2
from json import loads


def logmod(num: int):
    if num <=0:
        return []
    ab = []
    while True:
        a = int(log2(num))
        ab.append(a + 1)
        num = num - 2 ** a
        print(num)
        if num == 0:
            break
    return ab

def mark_bgi_ml_image():
    font = cv2.FONT_HERSHEY_SIMPLEX
    for kk, unit_folder_path in enumerate(data_pardir.iterdir()):
        print(f"--开始检查第 {kk} 個機臺{unit_folder_path.name} ML 信息")
        for ml_response_path in file_name(unit_folder_path, r'.*(SOBBK|ICEBK|BGI)\.txt'):
            ml_type = re.findall(r'.*(SOBBK|ICEBK|BGI)\.txt', ml_response_path.name)[0]
            pic_path = ml_response_path.parent / (ml_response_path.stem + '.JPG')
            
            print(f"----第 {kk} 個機臺{ml_type}信息檢查开始")
            with open(ml_response_path, "r", encoding='utf-8') as f:
                content = f.read()
            ml_response_status = re.findall(r'^HTTP/1.1 (\d+) (.*)$', content, re.M)
            if ml_response_status:
                status_code, response_message = ml_response_status[0]
            else:
                status_code, response_message = ('', '')
            print('----', status_code, response_message)
            
            result = re.findall(r'^\{.*}$', content, re.M)
            if result:
                ml_result = loads(result[0])
            else:
                ml_result = {}
            ml_pass_fail = ml_result.get('decision', status_code)
            pic_new_name = f'{ml_type}_{"_".join(pic_path.stem.split('.')[:3])}_{ml_pass_fail}'
            if ml_pass_fail == 0:
                print('------', f"該機臺{ml_type}無異常")
                pic_path = ml_response_path.parent / (ml_response_path.stem + '.JPG')
                img = cv2.imread(str(pic_path))
                cv2.putText(img, pic_new_name, (200, 100), font, 3, (255, 0, 0), 5)
                cv2.imwrite(str(pic_path.parent / f'{pic_new_name}{pic_path.suffix}'), img)
            else:
                print('------', ml_result)
                boxes = ml_result.get('bboxes', [])
                issue_flags = logmod(ml_result.get('issue_flags',0))
                roi_flags = logmod(ml_result.get('roi_flags',0))
                print(issue_flags,roi_flags)              
                img = cv2.imread(str(pic_path))
                for box in boxes:
                    hh, ww, xx, yy = box['bbox'].values()
                    cv2.rectangle(img, (xx, yy), (xx + ww, yy + hh), (0, 0, 255), 2)
                
                    cv2.putText(img, box['class_name'], (xx - 25, yy - 25), font, 1.5, (0, 0, 255), 3)
                cv2.putText(img, pic_new_name, (200, 100), font, 3, (255, 0, 0), 5)
                cv2.imwrite(str(pic_path.parent / f'{pic_new_name}{pic_path.suffix}'), img)
            print(f"----第 {kk} 個機臺{ml_type}信息檢查完成")
        print(f"--完成检查第 {kk} 個機臺{unit_folder_path.name} ML 信息\n")


if __name__ == "__main__":
    mark_bgi_ml_image()

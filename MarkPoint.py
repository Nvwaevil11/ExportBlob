# @File: MarkPoint.py
# @Time: 2024/1/28 下午 02:55
# @Author: Nan1_Chen
# @Mail: Nan1_Chen@pegatroncorp.com

# @Software: PyCharm
# --- --- --- --- --- --- --- --- ---

from math import log2
import re
from LoadFilePath import data_pardir, select_file, init_folders
from ExtractZIP import file_name, decompression_ml_image_txt
import cv2
from pathlib import Path
from json import loads
from pprint import pprint
import pandas as pd
from InsertImage import MLAlertTable

Issues = {
    1: "Flared Bat-Ear",
    2: "Open Side-Fold",
    3: "Exposed Matal",
    4: "Puncture",
    5: "Liquid",
    6: "Swelling",
    7: "Other",
    8: "Deformed Spring",
    9: "Foreign Material",
    10: "Extra Screw",
}

ml_types_dict = {
    "XGS": ["SOBBK", "ICEBK", "BGI", "SOB", "ICE"],
    "BGS1": ["SOBBK", "ICEBK", "BGI"],
    "CGS": ["SOB", "ICE"],
}

response_status = {
    200: "OK",
    400: "Bad request",
    403: "Forbidden",
    404: "Not found",
    405: "Method not allowed",
    413: "Payload too large",
    417: "Expectation Failed",
    422: "Unprocessable entity",
    500: "Internal server error",
    503: "Service unavailable",
    520: "Internal model error",
    1:"ML Fail",
    0: "ML Pass",
    -1: "ML Error",
    -2: "Unsuccessful Response",
    -3: "ML Request Timeout",
    -4: "Connection Error"
}

reasons = {
    1:"ML Fail",
    0: "ML Pass",
    -1: "ML Error",
    -2: "Unsuccessful Response",
    -3: "ML Request Timeout",
    -4: "Connection Error",
}
ml_file_suffixes = ["txt", "JPG"]


# 16进制颜色格式颜色转换为RGB格式
def hex_to_rgb(hex_color_code: str, reverse: bool = False) -> tuple:
    if hex_color_code.startswith("#"):
        r = int(hex_color_code[1:3], 16)
        g = int(hex_color_code[3:5], 16)
        b = int(hex_color_code[5:7], 16)
    else:
        r = int(hex_color_code[0:2], 16)
        g = int(hex_color_code[2:4], 16)
        b = int(hex_color_code[4:6], 16)
    if reverse:
        return b, g, r
    else:
        return r, g, b


def get_alert_list() -> dict:
    from LoadFilePath import alert_file_path

    if alert_file_path.suffix == ".csv":
        df = pd.read_csv(alert_file_path, parse_dates=["uut_start", "uut_stop"])
        df.set_index(["serial_number", "test_result", "uut_start"], inplace=True)
        df["unit_folder_name"] = [
            f'{_[0]}_{_[1]}_{_[2].strftime("%Y%m%d%H%M%S")}' for _ in df.index
        ]
        df.reset_index(inplace=True)
        df.set_index(["unit_folder_name"], inplace=True)
        df = df.convert_dtypes()
        for loc, col in enumerate(df.columns):
            if col.endswith("response_code"):
                df.insert(
                    loc,
                    col.replace("response_code", "response_status"),
                    df[col].apply(lambda x: response_status.get(x, x)),
                )
            if col.endswith("_decision"):
                df.insert(
                    loc,
                    col.replace("_decision", "_reason"),
                    df[col].apply(lambda x: reasons.get(x, x)),
                )

        return df.transpose().to_dict()
    else:
        raise FileNotFoundError("alert_file文件丢失")


class UnitFolder(object):
    def __init__(self, unit_folder_path: Path, ml_info: dict):
        self.test_data = ml_info
        self._station_name = ml_info.get("display_name", "XGS")
        self._ml_types = ml_types_dict.get(self._station_name, ml_types_dict["XGS"])
        self._ml_pattern = (
            rf'.*\.({"|".join(self._ml_types)})\.({"|".join(ml_file_suffixes)})$'
        )
        if unit_folder_path.exists():
            if unit_folder_path.is_dir():
                self._folder_path = unit_folder_path
            elif unit_folder_path.is_file():
                print(
                    f"Waring: 初始化路径[{unit_folder_path}]是文件,不是目录,强制转换为文件的父目录[unit_folder_path.parent]"
                )
                self._folder_path = unit_folder_path.parent
            else:
                raise TypeError(
                    f"Error:初始化失败,初始化路径[{unit_folder_path}]类型错误,应该为目录或者文件."
                )
        else:
            print(f"{unit_folder_path}不存在,創建空目錄防止報錯.")
            unit_folder_path.mkdir(parents=True, exist_ok=True)
            self._folder_path = unit_folder_path

    @property
    def unit_folder_path(self):
        return self._folder_path

    @property
    def station_name(self):
        return self._station_name

    def parse_ml_folder(self) -> dict:
        ml_files = file_name(self._folder_path, self._ml_pattern)
        if ml_files:
            for ml_file in ml_files:
                ml_type, ml_suffix = re.findall(self._ml_pattern, ml_file.name)[0]
                self.test_data.update(
                    {
                        f'{ml_type[:3].lower()}_{ml_suffix.replace("JPG", "pic")}': str(
                            ml_file
                        )
                    }
                )
        else:
            self.test_data = self.test_data
        return self.test_data


def parse_ml_folder(
    unit_folder: Path, ml_types: list|None = None, file_suffixes: list|None = None
) -> dict:
    if ml_types is None:
        ml_types = ["SOBBK", "ICEBK", "BGI", "SOB", "ICE"]
    if file_suffixes is None:
        file_suffixes = ["txt", "JPG"]
    if unit_folder.is_file():
        unit_folder = unit_folder.parent
    ml_folder_files = {}
    for ml_type in ml_types:
        ml_files = file_name(
            unit_folder, rf'.*\.{ml_type}\.({"|".join(file_suffixes)})$'
        )
        if ml_files:
            ml_folder_files.update(
                {ml_type[:3]: {_.suffix[1:]: str(_) for _ in ml_files}}
            )
        else:
            ml_folder_files.update({ml_type[:3]: {}})
    return {unit_folder.name: ml_folder_files}


def logmod(num: int):
    if num <= 0: 
        return []
    ab = []
    while True:
        a = int(log2(num))
        ab.append(a + 1)
        num = num - 2**a
        # print(num)
        if num == 0:
            break
    return ab


def mark_bgi_ml_image():
    font = cv2.FONT_HERSHEY_SIMPLEX
    from LoadFilePath import ml_station

    ml_types = ml_types_dict.get(ml_station, ml_types_dict["XGS"])
    for kkk, unit_folder_path in enumerate(data_pardir.iterdir()):
        print(f"--开始检查第 {kkk + 1} 個機臺{unit_folder_path.name} ML 信息")
        pprint(parse_ml_folder(unit_folder_path, ml_types=ml_types))
        for ml_response_path in file_name(
            unit_folder_path, r".*(SOBBK|ICEBK|BGI|SOB|ICE)\.txt"
        ):
            ml_type = re.findall(
                r".*(SOBBK|ICEBK|BGI|SOB|ICE)\.txt", ml_response_path.name
            )[0]
            pic_path = ml_response_path.parent / (ml_response_path.stem + ".JPG")
            if not pic_path.exists():
                pic_path = file_name(unit_folder_path, rf".*{ml_type}\.JPG")[0]
            print(f"----第 {kkk + 1} 個機臺{ml_type}信息檢查开始{pic_path.exists()}")
            with open(ml_response_path, "r", encoding="utf-8") as f:
                content = f.read()
            ml_response_status = re.findall(r"^HTTP/1.1 (\d+) (.*)$", content, re.M)
            if ml_response_status:
                status_code, response_message = ml_response_status[0]
                print("------", "ml_response_status:", [status_code, response_message])
            else:
                status_code, response_message = ("404", "not_found_response_status")
                print("------", "ml_response_status:", [status_code, response_message])

            result = re.findall(r"^\{.*}$", content, re.M)
            if result:
                ml_result = loads(result[0])
            else:
                ml_result = {}
            ml_pass_fail = ml_result.get("decision", -2)
            pic_new_name = (
                f'{ml_type}_{"_".join(pic_path.stem.split('.')[:3])}_'
                f'{ml_pass_fail}_{status_code}({response_message})'
            )
            target_pic_path = str(pic_path.parent / f"{pic_new_name}{pic_path.suffix}")
            print("------", f"目标图片路径{target_pic_path}")
            if ml_pass_fail == 0:
                print("------", f"該機臺{ml_type}無異常")
                img = cv2.imread(str(pic_path))
                cv2.putText(
                    img,
                    pic_new_name,
                    (200, 100),
                    font,
                    3,
                    hex_to_rgb("#0000FF", True),
                    5,
                )
                cv2.imwrite(target_pic_path, img)
            else:
                print("------", ml_result)
                boxes = ml_result.get("bboxes", [])
                issue_flags = logmod(ml_result.get("issue_flags", 0))
                roi_flags = logmod(ml_result.get("roi_flags", 0))
                print([Issues.get(_, "") for _ in issue_flags], roi_flags)
                img = cv2.imread(str(pic_path))
                for box in boxes:
                    hh, ww, xx, yy = box["bbox"].values()
                    cv2.rectangle(
                        img,
                        (xx, yy),
                        (xx + ww, yy + hh),
                        hex_to_rgb("#FF0000", True),
                        2,
                    )
                    cv2.putText(
                        img,
                        box["class_name"],
                        (xx - 25, yy - 25),
                        font,
                        1.5,
                        hex_to_rgb("#FF0000", True),
                        3,
                    )
                cv2.putText(
                    img,
                    pic_new_name,
                    (200, 100),
                    font,
                    3,
                    hex_to_rgb("#ff0000", True),
                    5,
                )
                cv2.imwrite(target_pic_path, img)
            print(f"----第 {kkk + 1} 個機臺{ml_type}信息檢查完成")
        print(f"--完成检查第 {kkk + 1} 個機臺{unit_folder_path.name} ML 信息\n")


if __name__ == "__main__":
    init_folders()
    select_file()
    decompression_ml_image_txt()
    alert_list = get_alert_list()
    pprint(alert_list)
    new_list = []
    for test_folder_path, test_data in alert_list.items():
        test_folder = UnitFolder(data_pardir / test_folder_path, test_data)
        new_list.append(test_folder.parse_ml_folder())
    pprint(new_list)
    df1 = pd.DataFrame(data=new_list)
    print(df1)
    df1.rename(
        columns={
            "model_sob_decision": "sob_decision",
            "model_ice_decision": "ice_decision",
            "model_bgi_decision": "bgi_decision",
            "model_ice_extra_screw_decision": "extra_screw_decision",
            "model_sob_reason": "sob_reason",
            "model_ice_reason": "ice_reason",
            "model_bgi_reason": "bgi_reason",
            "model_ice_extra_screw_reason": "extra_screw_reason",
            "model_sob_no_retest": "sob_no_retest",
            "model_sob_response_status": "sob_status",
            "model_ice_no_retest": "ice_no_retest",
            "model_ice_response_status": "ice_status",
            "model_bgi_response_status": "bgi_status",
            "model_bgi_no_retest": "bgi_no_retest"           
        },
        inplace=True,
    )
    cols = list(df1.columns)
    new_cols = [
        "serial_number",
        "uut_start",
        "display_name",
        "station_id",
        "sob_decision",
        "ice_decision",
        "bgi_decision",
        "sob_reason",
        "ice_reason",
        "bgi_reason",
        "bgi_pic",
        "ice_pic",       
        "sob_pic",
        "bgi_txt",
        "ice_txt",
        "sob_txt",
    ]
    df1.index = pd.Index(name='No.',data=range(1,len(df1)+1))
    df1.to_csv(r"D:\Users\Xiaoze_Wang\Desktop\BGS ML Alert SN re-judge tracker list2.csv",index_label='No.')
    df1 = df1[[_ for _ in new_cols if _ in cols]]
    add_image = MLAlertTable(test_info=df1)

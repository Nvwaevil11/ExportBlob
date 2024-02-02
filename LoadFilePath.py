# @File: LoadFilePath.py
# @Time: 2024/1/28 上午 10:12
# @Author: Nan1_Chen
# @Mail: Nan1_Chen@pegatroncorp.com


import shutil
import tkinter as tk
from tkinter import filedialog
from pathlib import Path
import datetime
from pprint import pprint

root_path: Path = Path(__file__).resolve().parent
load_dir: Path = root_path / "LoadFile"
alert_fail_units_dir: Path = load_dir / "AlertFailUnits"
alert_file_path = alert_fail_units_dir
data_dir: Path = load_dir / "DataZip"
par_dir: Path = root_path / "ParseFile"
alert_fail_units_pardir: Path = par_dir / "AlertFailUnits"
data_pardir: Path = par_dir / "DataZip/Data"
data_zippardir: Path = par_dir / "DataZip/DataZip"
export_dir: Path = root_path / 'ExportFile'

ml_station = 'XGS'


def function_timer(func):  # 計時器
    def inner(*args, **kwargs):
        start = datetime.datetime.now()
        print(f'{datetime.datetime.now()} Function "{
              func.__name__}" start running -->')
        result = func(*args, **kwargs)
        end = datetime.datetime.now()
        _ = (end - start).total_seconds()
        print(f'{datetime.datetime.now()} Function "{
              func.__name__}" finished running in {_:.6f}s<--')
        return result

    return inner


def init_folders():
    print(f'---1.初始化根级目录', end='....->')
    shutil.rmtree(load_dir, ignore_errors=True)
    shutil.rmtree(par_dir, ignore_errors=True)
    if not export_dir.exists():
        export_dir.mkdir(parents=True, exist_ok=True)
    print(f'Done!!')

    print(f'---2.创建子级目录', end='....->')
    alert_fail_units_dir.mkdir(parents=True, exist_ok=True)
    data_dir.mkdir(parents=True, exist_ok=True)
    alert_fail_units_pardir.mkdir(parents=True, exist_ok=True)
    data_pardir.mkdir(parents=True, exist_ok=True)
    data_zippardir.mkdir(parents=True, exist_ok=True)
    print(f'Done!!')


def select_file():
    global alert_file_path
    global ml_station
    root = tk.Tk()
    root.withdraw()  # 隐藏主窗口
    print(f"---3.选择AlertFailUnits文件，", end='....->')
    file_path = filedialog.askopenfilename(title='请选择AlertFailUnits文件',
                                           filetypes=[('cgsAlertFailUnits文件', '*_ml_alert_*_units_ae34_*.csv;*_ml_alert_*_units_station1614_*.csv')], )
    if file_path != '':
        shutil.copy(file_path, alert_fail_units_dir)
        alert_file_path = alert_fail_units_dir / Path(file_path).name
        if '_ae34_' in file_path:
            ml_station = 'CGS'
        elif '_station1614_' in file_path:
            ml_station = 'BGS'
        print(f"Done!!")
        print(f'-----复制选择的文件[{file_path}] 到 文件夹[{alert_fail_units_dir}].')
    else:
        print(f'Error!!  --不可以取消操作,会退出哦..')
        raise FileNotFoundError('用户选择了取消,结束进程...')
    print(f"---4.选择Bolb压缩包文件(可多选)，", end='....->')
    file_path1 = filedialog.askopenfilenames(title='请选择Insight下载的Bolb压缩包文件',
                                             filetypes=[('Bolb压缩包文件', 'data*.zip')])

    if isinstance(file_path1, tuple):
        print('Done!!')
        for zip_data in file_path1:
            shutil.copy(zip_data, data_dir)
            print(f'-----复制选择的文件[{zip_data}] 到 文件夹[{data_dir}].')
    else:
        print(f'Error!!  --不可以取消操作,会退出哦..')
        raise FileNotFoundError('用户选择了取消,结束进程...')


if __name__ == '__main__':
    select_file()

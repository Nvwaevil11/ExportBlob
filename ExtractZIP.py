# @File: ExtractZIP.py
# @Time: 2024/1/28 下午 01:39  
# @Author: Nan1_Chen
# @Mail: Nan1_Chen@pegatroncorp.com

# @Software: PyCharm
# --- --- --- --- --- --- --- --- ---
import shutil
import tarfile
from zipfile import ZipFile
from pathlib import Path
import re
from LoadFilePath import data_dir, data_zippardir, data_pardir

root_path = Path(__file__).resolve().parent


def unzip2folder(filename: Path, out_path: Path = None):
    if not out_path:
        out_path = filename.parent
    with ZipFile(filename) as zip1:
        for file in zip1.filelist:
            if not file.is_dir():
                with zip1.open(file) as source:
                    with open(out_path / file.filename.split("/")[-1], "wb") as target_file:
                        shutil.copyfileobj(source, target_file)


def untgz2folder(filename: Path, out_path: Path = None, filter_str: str = None):
    member: tarfile.TarInfo
    if not out_path:
        out_path = filename.parent
    tar1 = tarfile.open(filename)

    with tar1:
        members = tar1.getmembers()
        root_folder = Path(members[0].path).parts[0]
        target_folder = out_path / root_folder
        for member in tar1.getmembers():
            if not member.isdir():
                if filter_str:
                    if not re.findall(filter_str, member.name):
                        continue
                target_folder.mkdir(parents=True, exist_ok=True)
                source = tar1.extractfile(member)
                target_file = open(target_folder / Path(member.name).name, 'wb')
                with source, target_file:
                    shutil.copyfileobj(source, target_file)
    return


def file_name(
        file_dir: Path, file_type: str
) -> list:  # 獲取目錄中擴展名為file_type的文件列表
    """
    获取指定目录下所有指定类型的文件的列表
    :param file_dir:文件目录
    :param file_type: 文件类型,支持正則表達式
    :return: 文件列表
    """
    ls = []
    files = file_dir.glob("**/*")

    for file in files:
        # print(file_type,re.search(file_type,file.name),file.name)
        if re.search(file_type, file.name):
            ls.append(file)
    return ls


def decompression_ml_image_txt():
    print("---5.文件正在解壓 zip -> tgz...")
    for zip_file_index, zip_file in enumerate(file_name(data_dir, r'data.*\.zip')):
        unzip2folder(zip_file, data_zippardir)
        print(f'-----解压第{zip_file_index + 1}个压缩包[{zip_file}],', end='...->')
        zip_file.unlink()
        print('Done!!')
    print("---Done!!")
    print("---6.文件已完成第一步解壓: tgz ...")
    [_.unlink() for _ in file_name(data_zippardir, r'.*_blob777\.tgz$')]
    for tgz_file in file_name(data_zippardir, r'.*_log\.tgz$'):
        untgz2folder(tgz_file, data_pardir, r'.*_FAIL_.*(SOBBK|ICEBK|BGI|SOB|ICE)\.(JPG|txt)$')
        tgz_file.unlink()

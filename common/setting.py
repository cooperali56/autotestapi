import os
from typing import Text


def root_path() -> Text:
    """ 获取 根路径 """
    path = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    return path


def ensure_path_sep(path: Text) -> Text:
    """兼容 windows 和 linux 不同环境的操作系统路径 """
    if "/" in path:
        path = os.sep.join(path.split("/"))
    
    if "\\" in path:
        path = os.sep.join(path.split("\\"))
    
    return root_path() + path


def find_file_path(filename):
    """
    查找文件路径
    :param filename: 带后缀文件名
    :return:
    """
    for root, dirs, files in os.walk(root_path()):
        for file in files:
            if file == filename:
                return os.path.join(root, file)
    raise FileNotFoundError(f"File '{filename}' not found in project")


if __name__ == '__main__':
    try:
        file_path = find_file_path('kol.xlsx')
        print(f"File found at: {file_path}")
    except FileNotFoundError as e:
        print(e)

    print(ensure_path_sep('kol.xlsx'))

import json

from utils.logging_tool.log_control import ERROR


def find_value(data_dict, target_key):
    """
    查找嵌套数据
    :param data_dict:
    :param target_key:
    :return:
    """
    if target_key in data_dict:
        return data_dict[target_key]
    elif isinstance(data_dict, dict):
        for key, value in data_dict.items():
            if isinstance(value, dict):
                result = find_value(value, target_key)
                if result is not None:
                    return result
            elif isinstance(value, list or tuple):
                for item in value:
                    result = find_value(item, target_key)
                    if result is not None:
                        return result
    return None


def get_url_path(url) -> str:
    """
    从url中获取path
    :param url:
    :return:
    """
    com_index = url.find("com")
    if com_index != -1:
        path = url[com_index + len("com"):]
        return path
    return url


def get_value_by_path(json_data, path) -> object:
    """
    获取 JSON 数据中指定路径的值
    :param json_data: 输入的 JSON 数据
    :param path: 切片路径，以列表形式提供
    :return: 如果找到路径，返回对应的值；否则返回 None
    """
    for key in path:
        if isinstance(json_data, list):
            try:
                key = int(key)
            except ValueError:
                return None
        if key in json_data:
            json_data = json_data[key]
        else:
            return None
    return json_data


def set_data_key(data) -> list:
    """
    映射测试用例key
    :param data:
    :return:
    """

    # 定义键名映射
    key_mapping = {
        '用例title': '用例标题',
        '请求path': 'url',
        '请求方式': 'method',
        '请求头headers': 'headers',
        '请求data': 'data',
        '数据type': 'data_type',
    }
    # 创建新的数据
    new_data = []

    # 使用列表推导式和字典的get方法创建新的数据
    new_data = [{key_mapping.get(key, key): value for key, value in item.items()} for item in data]

    return new_data


def handle_excel_dict(data_excel) -> list:
    """
    内部处理Excel数据为指定字典格式
    :param data_excel:
    :return:
     """
    data_excel_title = data_excel[0]
    data_excel_test = []
    for i in range(1, len(data_excel)):
        data_excel_dict = {}
        for j, title in enumerate(data_excel_title):
            data_excel_dict[title] = data_excel[i][j]
        data_excel_test.append(data_excel_dict)
    return data_excel_test

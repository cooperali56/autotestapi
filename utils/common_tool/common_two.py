from jsonpath_ng import parse
from random import randint

global_variable = {}


def get_params(response, ex, key):
    # 创建 JSON 路径表达式
    expression = parse(ex)
    # 在数据中查找匹配的结果
    matches = [match.value for match in expression.find(response)]
    for match in matches:
        if match != '' or match is not None:
            global_variable[key] = match


def set_params(request_info):
    if global_variable.keys():
        new = ""
        for var in global_variable.keys():
            new = str(request_info).replace(f"{{{{{var}}}}}", global_variable[var])
        return eval(new)
    else:
        return request_info


# 随机10位字符串数字
# 可用于随机uname、email等参数
def random_string() -> str:
    number = ""
    for i in range(0, 10):
        number += str(randint(0, 10))
    return number

from numpy.core.defchararray import strip

from utils.common_tool.common_data import find_value


def assert_test(actual_data: dict, expect_data, assert_msg=None) -> bool:
    """
    断言方法
    :param actual_data: 实际数据
    :param expect_data: 预期数据
    :param assert_msg:  断言信息
    :return:
    """

    # 从 expect_data 中提取所需的信息
    expect_key = str(expect_data['key'])
    assert_type = str(expect_data['type'])
    expect_value = str(expect_data['value'])

    # 获取实际值
    actual_value = str(find_value(actual_data, expect_key))

    # 构建断言信息
    assert_msg_msg = f"断言信息=【{assert_msg}】；actual=【{actual_value}】 VS expect=【{expect_value}】"
    # 定义断言字典
    assert_dict = {
        '==': actual_value == expect_value,
        '!=': actual_value != expect_value,
        'in': actual_value in expect_value,
        'not in': actual_value not in expect_value,
        'is': actual_value is expect_value,
        'is not': actual_value is not expect_value,
    }

    # 捕获断言异常
    print("\n")
    try:
        assert assert_dict.get(assert_type, False), assert_msg_msg
    except AssertionError:
        raise AssertionError


def assert_all(actual, type_ass, expect, msg=None) -> bool:
    """
    断言
    :param actual:
    :param type_ass:
    :param expect:
    :param msg:
    :return:
    """
    if type_ass == '==':
        assert actual == expect, msg
    elif type_ass == '!=':
        assert actual != expect, msg
    elif type_ass == 'in':
        assert expect in actual, msg
    elif type_ass == 'not in':
        assert expect not in actual, msg
    elif type_ass == '>':
        assert actual > expect, msg
    elif type_ass == '>=':
        assert actual >= expect, msg
    elif type_ass == '<':
        assert actual < expect, msg
    elif type_ass == '<=':
        assert actual <= expect, msg
    else:
        return False

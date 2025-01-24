import re

import requests

from utils.logging_tool.log_control import *


def output_api(data: dict, api: requests, actual=True):
    """
    打印接口响应信息
    :param data:
    :param api:
    :param actual: 实际结果
    :return:
    """
    # 判断 actual 是否为 True
    if actual:
        print('\n')
        INFO.logger.info(f"\n======================================================\n"
                         f"用例ID: {data['API_CASE_ID']}\n"
                         f"用例标题: {data['用例title']}\n"
                         f"请求路径: {api.url}\n"
                         f"请求方式: {data['请求方式']}\n"
                         # f"请求头:   {api.headers}\n"
                         f"请求内容: {data['请求data']}\n"
                         f"接口响应状态码: {api.status_code}\n"
                         f"接口响应时长: {api.elapsed.total_seconds()}\n"
                         f"预期断言内容: {data['预期结果']}\n"
                         f"断言结果: {actual} = Assertion:Success!\n"
                         "=====================================================")
    else:
        print('\n')
        ERROR.logger.error(f"\n======================================================\n"
                           f"用例ID: {data['api-case-id']}\n"
                           f"用例标题: {data['用例title']}\n"
                           f"请求路径: {api.url}\n"
                           f"请求方式: {data['请求方式']}\n"
                           # f"请求头:   {api.headers}\n"
                           f"请求内容: {data['请求data']}\n"
                           f"接口响应状态码: {api.status_code}\n"
                           f"接口响应时长: {api.elapsed.total_seconds()}\n"
                           f"预期断言内容: {data['预期结果']}\n"
                           f"断言结果: {actual} = Assertion:Failed!\n"
                           "=====================================================")


def print_api_info(data: dict, api: requests, ass_status):
    """
    打印接口响应信息
    :param data:
    :param api:
    :param ass_status: (None=默认、0=成功、1=失败、2=跳过)
    :return:
    """

    log_info = {
        '项目模块': data['项目模块'],
        '用例ID': data['api-case-id'],
        '用例标题': data['用例标题'],
        '请求URL': api.url,
        '请求method': data['method'],
        '请求头headers': data['headers'],
        '请求cookies': data['cookies'],
        '请求数据data': data['data'],
        '接口响应状态码': api.status_code,
        '接口响应时长': api.elapsed.total_seconds(),
        '接口响应json': api.json(),
        '预期断言内容': data['预期结果']
    }

    log_message = f"\n======================================================\n"
    for key, value in log_info.items():
        log_message += f"{key} => {value}\n"
    print('\n')
    if ass_status == 0:
        log_message += "断言结果 => 成功!\n"
        INFO.logger.info(log_message)
    elif ass_status == 1:
        log_message += f"断言结果 => 失败!\n"
        ERROR.logger.error(log_message)
    elif ass_status == 2 or ass_status is None:
        log_message += "断言结果 => 预期为空不做断言处理!\n"
        DEBUG.logger.debug(log_message)

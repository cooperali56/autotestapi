import allure

from apiconfig.api_base import ApiBase
from utils.assert_tool.assert_type import assert_test
from utils.common_tool.common_data import find_value, get_url_path
from utils.logging_tool.log_decorator import print_api_info


def allure_set(data_test, data_yaml) -> None:
    """
    设置allure
    :param data_test:
    :param data_yaml:
    :return:
    """
    allure.dynamic.epic(f"项目名称: {data_yaml['name']}")
    allure.dynamic.feature(f"项目模块: {data_test['项目模块']}")
    allure.dynamic.title(f"用例标题：{data_test['用例标题']}")
    allure.dynamic.id(f"用例id：{data_test['api-case-id']}")
    allure.dynamic.story(f"接口url: {data_test['url']}")


def assert_api(api_base: ApiBase, data_yaml) -> None:
    """
    断言操作
    """
    # 默认为跳过
    ass_status = 2
    request = api_base.get_api_data()
    expected_result = request['预期结果']
    api_json = api_base.get_api().json()
    # 如果expected_result中的value带有$
    if '$' in expected_result['value']:
        expected_result['value'] = expected_result['value'].replace('$', '')
        expected_result['value'] = data_yaml['depend'][expected_result['value']]
    if expected_result is not None and expected_result['key'] != '':
        try:
            assert_test(api_json, expected_result, request['用例标题'])
            ass_status = 0
        except (KeyError, UnboundLocalError, AssertionError):
            ass_status = 1
            raise AssertionError
        finally:
            print_api_info(data=request, api=api_base.get_api(), ass_status=ass_status)
    else:
        print_api_info(data=request, api=api_base.get_api(), ass_status=ass_status)


def depend_config(data_test, data_yaml, api_base_json=None) -> dict:
    """
    依赖配置
    :param data_test:       测试用例
    :param data_yaml:       yaml数据
    :param api_base_json:   api响应json数据
    """
    # 依赖 - 赋值
    depend_data = data_test['依赖']
    depend_code: int = depend_data['code']
    depend_exp: list = depend_data['exp']
    depend_get: list = depend_data['get']
    depend_set: list = depend_data['set']
    depend_yaml: dict = data_yaml['depend']

    # 依赖 - 取
    if depend_code in (2, 3):
        # 获取多少个值
        get_data = []
        for get_s in depend_get:
            get_data.append(depend_yaml[get_s])
        # 遍历这个值的内容给依赖公式
        for i, get_s in enumerate(depend_exp):
            use = get_s['use']
            key = get_s['key']
            data_test[use][key] = get_data[i] if len(get_data) != 1 else get_data[0]

    # 依赖 - 存
    if api_base_json is not None:
        if depend_code in (1, 3):
            # 存数据源
            api_json = api_base_json
            # 遍历公式=list
            for set_kay in depend_set:
                if set_kay == 'all':
                    depend_yaml = api_json
                else:
                    depend_yaml[set_kay] = find_value(api_json, set_kay)
            return depend_yaml

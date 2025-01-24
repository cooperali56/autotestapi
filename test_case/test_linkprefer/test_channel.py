import allure
import pytest

from apiconfig.base_api import BaseApi
from utils.assert_tool.assert_type import assert_all
from utils.excel_tool.excel_control import ExcelControl
from utils.logging_tool.log_control import *
from utils.logging_tool.log_decorator import output_api
from utils.yaml_tool.yaml_control import YamlControl

# 获取excel测试数据
login_data_test = ExcelControl("data/linkprefer/linkprefer_channel.xlsx").get_excel_data("channel")
# 获取yaml配置
data_yaml = YamlControl("common/config.yaml").get_yaml_data()

@allure.epic(f"{data_yaml['linkprefer']['project_name']}")
@allure.feature(f"模块名称: {login_data_test[0]['项目模块']}")
class TestChannel:

    @allure.story(f"接口名称: {login_data_test[0]['请求path']}")
    @allure.title(f"用例标题: {login_data_test[0]['用例title']}")
    @pytest.mark.parametrize('in_data', login_data_test)
    def test_channel(self, in_data):

        # 组合url
        in_data['请求path'] = data_yaml['linkprefer']['test_host'] + in_data['请求path']

        # 发送请求
        api_http_object = BaseApi(**in_data)

        # case断言需求
        case_assert = eval(in_data['预期结果'])

        actual_assert = api_http_object.get_api_json()[case_assert['key']]

        # 日志
        output_api(in_data, api_http_object.get_api_object(), actual_assert)

        # 断言
        try:
            assert_all(actual_assert, case_assert['type'], case_assert['value'])
            INFO.logger.info("Assertion succeeded: Success!")

        except AssertionError:
            ERROR.logger.error(
                f"Assertion failed: \n"
                f"断言接口=【{api_http_object.get_api_url()}】失败了！！！"
                f"响应拿到的结果为【{actual_assert}】{case_assert['type']} 预期结果为【{case_assert['value']}】有误！！！\n"
                f"该接口的请求参数为 \n{in_data['请求data']}"
            )
            raise AssertionError


if __name__ == '__main__':
    pytest.main(['-q', 'test_channel.py'])

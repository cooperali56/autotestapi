import pytest

from apiconfig.api_object import ApiObject
from utils.common_tool.common_api import allure_set
from utils.common_tool.common_data import set_data_key
from utils.common_tool.common_two import random_string, get_params
from utils.excel_tool.excel_control import get_all_excel
from utils.yaml_tool.yaml_control import YamlControl


class TestAccount:

    def setup_class(self):
        """
        配置全局变量值
        """

        # 公共数据
        yaml_file = YamlControl('common/config.yaml').get_yaml_data()
        self.data_yaml = yaml_file['LinkBux']
        self.api = ApiObject()

    def teardown_class(self):
        """
        测试后清除
        """
        del self.api
        del self.data_yaml

    @pytest.mark.parametrize('data_test', set_data_key(get_all_excel(file_name='linkbux_account.xlsx', sheet='login')))
    def test_login(self, data_test):
        """
        测试登录
        """

        # 配置allure
        allure_set(data_test, self.data_yaml)

        # 组合url
        data_test['url'] = self.data_yaml['host'] + data_test['url']

        # api操作
        self.api.api_object_bast(data_test)

        # 依赖操作
        get_params(self.api.get_api_json(), data_test['依赖表达式'], data_test['依赖'])

    @pytest.mark.parametrize('data_test', set_data_key(get_all_excel(file_name='linkbux_account.xlsx', sheet='register')))
    def test_register(self, data_test):
        """
        测试注册
        """

        # 配置allure
        allure_set(data_test, self.data_yaml)

        # 组合url
        data_test['url'] = self.data_yaml['host'] + data_test['url']

        # 请求参数随机
        random_uname = "zhou" + random_string()
        random_email = random_string() + "@55haitao.com"
        data_test['data'] = data_test['data'].replace("{{uname}}", random_uname)
        data_test['data'] = data_test['data'].replace("{{email}}", random_email)

        # api操作
        self.api.api_object_bast(data_test)

        # 依赖操作
        get_params(self.api.get_api_json(), data_test['依赖表达式'], data_test['依赖'])

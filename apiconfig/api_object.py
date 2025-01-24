from apiconfig.api_base import ApiBase

from utils.common_tool.common_api import assert_api, depend_config, allure_set
from utils.common_tool.common_data import get_url_path
from utils.logging_tool.log_decorator import print_api_info
from utils.yaml_tool.yaml_control import YamlControl


class ApiObject:
    """
    一个接口对象封装
    """

    def __init__(self):
        self.api_base = ApiBase()

    def get_api_json(self) -> dict:
        """
        获取api_json
        :return:
        """
        return self.api_base.get_api().json()

    def _api_init(self, data_test, data_yaml=None):
        """
        初始化数据
        """
        self.api_base.set_api_data(data_test)
        self.api_base.api_data_to_dict('headers', 'cookies', 'data', '预期结果', '依赖')
        # 配置allure/组合数据
        allure_set(data_test, data_yaml)
        # 替换url 如果里面带有http就不替换
        if 'http' not in data_test['url']:
            data_test['url'] = data_yaml['host'] + get_url_path(data_test['url'])

    def api_object_bast(self, data_test) -> dict:
        """
        基础的接口
        :param data_test: 测试数据
        """
        # 初始化参数
        self._api_init(data_test, None)

        # 发送请求
        self.api_base.api_send_select()

        # 打印
        print_api_info(data=data_test, api=self.api_base.get_api(), ass_status=None)

        # 返回结果
        return self.get_api_json()

    def api_object_depend(self, data_test, yaml_obj: YamlControl) -> None:
        """
        接口依赖场景
        :param data_test: 测试数据
        :param yaml_obj: yaml对象
        """

        # yaml数据
        data_yaml = yaml_obj.get_yaml_data()

        # 初始数据
        self._api_init(data_test, data_yaml)

        # 依赖 - 取
        depend_config(data_test, data_yaml)

        # 发送请求
        self.api_base.api_send_select()

        # 依赖 - 存
        depend_yaml = depend_config(data_test, data_yaml, self.get_api_json())

        # 断言
        assert_api(self.api_base, data_yaml)

        # 存yaml
        if depend_yaml is not None:
            yaml_obj.update_yaml_value('depend', depend_yaml)

    def updata(self):
        pass
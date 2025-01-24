import pytest

from apiconfig.api_object import ApiObject
from utils.excel_tool.excel_control import get_all_excel
from utils.yaml_tool.yaml_control import YamlControl


class TestDemoBase:

    def setup_class(self):
        """
        配置全局变量值
        """
        # 公共数据
        yaml_file = YamlControl('demo_base.yaml').get_yaml_data()
        self.data_yaml = yaml_file['test_demo']
        self.api = ApiObject()

    #
    def teardown_class(self):
        """
        测试后清除
        """
        del self.api

    @pytest.mark.parametrize('data_test', get_all_excel(file_name='demo_base.xlsx', sheet='sheet_demo'))
    def test_demo_01(self, data_test):
        """
        测试case01
        """
        # api操作
        self.api.api_object_bast(data_test)


if __name__ == '__main__':
    pytest.main(['-q', 'test_demo_base.py'])

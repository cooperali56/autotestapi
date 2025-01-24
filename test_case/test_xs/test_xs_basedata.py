import pytest

from apiconfig.api_object import ApiObject
from test_case.test_kolsaas import yaml_name
from utils.excel_tool.excel_control import get_all_excel
from utils.yaml_tool.yaml_control import YamlControl


class TestXSBaseData:

    def setup_class(self):
        """
        测试前声明
        """
        # 公共数据
        self.yaml_object = YamlControl("xs.yaml")
        self.api = ApiObject()

    #
    def teardown_class(self):
        """
        测试后清除
        """
        del self.api
        del self.yaml_object

    @pytest.mark.parametrize('data_test', get_all_excel(file_name='xs_basedata.xlsx', sheet='BU列表'))
    # @pytest.mark.run(order=1)
    def test_bu_list(self, data_test):
        self.api.api_object_depend(data_test, self.yaml_object)




if __name__ == '__main__':
    pytest.main(['-q', 'test_xs_basedata.py'])

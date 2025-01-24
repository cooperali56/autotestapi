import pytest

from apiconfig.api_object import ApiObject
from utils.excel_tool.excel_control import get_all_excel
from utils.yaml_tool.yaml_control import YamlControl


class TestKolList:

    def setup_class(self):
        """
        配置全局变量值
        """
        # 公共数据
        self.yaml_object = YamlControl('kolsaas.yaml')
        self.api = ApiObject()

    #
    def teardown_class(self):
        """
        测试后清除
        """
        del self.api
        del self.yaml_object

    @pytest.mark.parametrize('data_test', get_all_excel(file_name='KOLList.xlsx', sheet='showKolList'))
    def test_showKolList(self, data_test):
        self.api.api_object_depend(data_test, self.yaml_object)


if __name__ == '__main__':
    pytest.main(['-q', 'test_kollist.py'])

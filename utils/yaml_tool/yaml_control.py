import os
from ruamel.yaml import YAML
from common.setting import find_file_path, ensure_path_sep


class YamlControl:
    """
    获取 YAML 文件中的数据
    """

    def __init__(self, file_name):
        """
        实例化类属性
        :param file_name: 文件名；带上前路径
        """
        self.file_name = file_name
        self.__yaml_path = None

    def _load_yaml(self):
        """
        内部方法：加载 YAML 文件
        :return: 解析后的 YAML 数据
        """
        try:
            with open(self.__yaml_path, 'r', encoding='utf-8') as file:
                yaml = YAML(typ='safe')
                return yaml.load(file)
        except (FileNotFoundError, yaml.YAMLError) as e:
            raise Exception(f"Failed to load YAML file: {self.file_name}") from e

    def get_yaml_data(self):
        """
        获取 YAML 文件中的数据
        :return: 解析后的 YAML 数据
        """
        if os.path.sep in self.file_name:
            self.__yaml_path = ensure_path_sep(f"\\{self.file_name}")
            return self._load_yaml()
        else:
            self.__yaml_path = find_file_path(self.file_name)
            return self._load_yaml()

    def update_yaml_value(self, key_path, new_value):
        """
        更新 YAML 文件中的值
        :param key_path: 键路径如 'auth'
        :param new_value: 新值
        """
        self.__yaml_path = find_file_path(self.file_name)
        yaml = YAML()
        try:
            with open(self.__yaml_path, 'r') as file:
                data = yaml.load(file)

            def update_recursively(node, keys, new_value):
                key = keys.pop(0)
                if not keys:
                    node[key] = new_value
                else:
                    update_recursively(node[key], keys, new_value)

            keys = key_path.split('-')  # 按照分隔符切割键路径
            update_recursively(data, keys, new_value)

            # 保存修改后的 YAML 文件
            with open(self.__yaml_path, 'w') as file:
                yaml.dump(data, file)
        except (FileNotFoundError, yaml.YAMLError, KeyError, IndexError) as e:
            raise Exception(f"Failed to update YAML file: {self.file_name}") from e

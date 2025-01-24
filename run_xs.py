import os

import pytest

from utils.logging_tool.log_control import INFO
from utils.yaml_tool.yaml_control import YamlControl

data_yaml = YamlControl('xs.yaml').get_yaml_data()


def run_xs():
    INFO.logger.info(
        """
  ____ ____  _           _ _                _   _   _ _____ ___ _____ _____ ____ _____
 | ___| ___|| |__   __ _(_) |_ __ _  ___   / \ | | | |_   _/ _ \_   _| ____/ ___|_   _|
 |___ \___ \| '_ \ / _` | | __/ _` |/ _ \ / _ \| | | | | || | | || | |  _| \___ \ | |
  ___) |__) | | | | (_| | | || (_| | (_) / ___ \ |_| | | || |_| || | | |___ ___) || |
 |____/____/|_| |_|\__,_|_|\__\__,_|\___/_/   \_\___/  |_| \___/ |_| |_____|____/ |_|
开始执行{}项目,测试人员{}...
""".format(data_yaml['name'], data_yaml['testers'])
    )

    pytest.main(['-s', '-W', 'ignore:Module already imported:pytest.PytestWarning',
                 '--alluredir', './report/tmp', "--clean-alluredir", "test_case/test_xs/"])

    """
        --reruns: 失败重跑次数
           --count: 重复执行次数
           -v: 显示错误位置以及错误的详细信息
           -s: 等价于 pytest --capture=no 可以捕获print函数的输出
           -q: 简化输出信息
           -m: 运行指定标签的测试用例
           -x: 一旦错误，则停止运行
           --maxfail: 设置最大失败次数，当超出这个阈值时，则不会在执行测试用例
            "--reruns=3", "--reruns-delay=2"
           """
    os.system(r"allure generate ./report/tmp -o ./report/html --clean")

    os.system(f"allure serve ./report/tmp -h 127.0.0.1 -p 9990")


if __name__ == '__main__':
    run_xs()

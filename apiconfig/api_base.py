import json
import time
from typing import Any

import requests

from apiconfig.api_set_data import ApiSetData
from apiconfig.api_set_send import ApiSetSend
from utils.logging_tool.log_control import ERROR


class ApiBase(ApiSetData, ApiSetSend):
    """
    接口基类
    """

    def _delete_send(self):
        print("asd")
        pass

    def __init__(self):
        self.__object = None
        self.__api_data = None
        requests.DEFAULT_RETRIES = 5
        self.__api_session = requests.session()

    def get_api_data(self) -> dict:
        """
        返回接口数据
        """
        return self.__api_data

    def set_api_data(self, data) -> None:
        """
        设置接口数据
        """
        self.__api_data = data

    def get_api(self) -> requests.Request:
        """
        获取接口
        """
        return self.__object

    def _get_send(self):
        """
        发送get请求
        """
        params = self.__api_data['data']
        return self.__api_session.get(
            url=self.__api_data['url'],
            headers=self.__api_data['headers'],
            params=params,
        )

    def _post_send(self):
        """
        发送post请求
        """
        data_type = self.__api_data['data_type']
        if data_type == 'form-data':
            return self.__api_session.post(
                url=self.__api_data['url'],
                headers=self.__api_data['headers'],
                data=self.__api_data['data'],
            )
        else:
            return self.__api_session.post(
                url=self.__api_data['url'],
                headers=self.__api_data['headers'],
                json=self.__api_data['data'],
            )

    def api_data_to_dict(self, *args):
        """
        将接口数据转化成字典
        """
        for arg in args:
            arg_data = self.__api_data[arg]
            try:
                if isinstance(arg_data, str):
                    arg_data = json.dumps(arg_data, indent=4)
                    arg_data = json.loads(arg_data)
                    self.__api_data[arg] = arg_data
                    self.__api_data[arg] = eval(arg_data)
                else:
                    self.__api_data[arg] = eval(arg_data)
            except (json.JSONDecodeError, SyntaxError, AttributeError, TypeError) as e:
                ERROR.logger.error(f"解析JSON或求值表达式失败: 解析名称: {arg}. 内容: {arg_data}")

    def api_send_select(self) -> Any:
        """
        发送请求
        """
        method = self.__api_data['method'].lower()
        self.__api_session.verify = False
        self.__api_data['headers']['Connection'] = 'close'
        time.sleep(0.2)
        if method == 'get':
            self.__object = self._get_send()
        elif method == 'post':
            self.__object = self._post_send()
        else:
            ERROR.logger.error(f"请求方法错误：{self.__api_data['method']}")
        return self.__object



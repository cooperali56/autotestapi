import requests

from apiconfig.set_request import *


class BaseApi:
    """ http基类"""

    # 接口对象
    __api_object: requests
    __api_response: requests.Response

    def __init__(self, **kwargs):
        """
        初始化
        :param kwargs: 请求参数
        """
        self.__api_object = send_api(**kwargs)
        self.__api_response = self.__api_object

    def get_api_object(self):
        """
        获取api对象
        :return:
        """
        return self.__api_object

    def get_api_json(self) -> dict:
        """
        获取api_response,json
        :return:
        """
        return self.__api_response.json()

    def get_api_url(self):
        """
        获取api-url
        :return:
        """
        return self.__api_response.url

    def get_api_text(self):
        """
        获取api_response,text
        :return:
        """
        return self.__api_response.text

import requests

from utils.common_tool.common_two import set_params, get_params
from utils.excel_tool.excel_control import ExcelControl
from utils.yaml_tool.yaml_control import YamlControl


def send_api(**kwargs) -> requests:
    """
    设置发送请求
    :param kwargs:
    :return:
    """

    # 格式转换
    kwargs['请求头headers'] = eval(kwargs['请求头headers'])
    kwargs['请求data'] = eval(kwargs['请求data'])

    if kwargs['请求方式'] == 'get':
        new_kwargs = set_params(kwargs)
        response = request_get(url_get=kwargs['请求path'], params_get=new_kwargs['请求data'],
                               headers=new_kwargs['请求头headers'])
        get_params(response.json(), kwargs['依赖表达式'], kwargs['依赖'])
        return response

    elif kwargs['请求方式'] == 'post':
        new_kwargs = set_params(kwargs)
        response = request_post(url_post=kwargs['请求path'], headers_post=new_kwargs['请求头headers'],
                                data_post=new_kwargs['请求data'])
        get_params(response.json(), kwargs['依赖表达式'], kwargs['依赖'])
        return response

    elif kwargs['请求方式'] == 'put':
        response = request_put(url_put=kwargs['请求path'], headers_put=kwargs['请求头headers'],
                               data_put=kwargs['请求data'])
        return response

    elif kwargs['请求方式'] == 'delete':
        response = request_delete(url_p=kwargs['请求path'], headers_p=kwargs['请求头headers'],
                                  data_p=kwargs['请求data'])
        return response

    else:
        raise Exception('请求方式错误')


def request_get(url_get, params_get, **kwargs) -> requests:
    """
    http之get请求
    :param url_get:
    :param params_get:
    :return:
    """
    return requests.get(url=url_get, params=params_get, **kwargs)


def request_post(url_post, headers_post, data_post, **kwargs) -> requests:
    """
    http之post请求
    :param url_post:
    :param headers_post:
    :param data_post:
    :return:
    """
    return requests.post(url=url_post, headers=headers_post, data=data_post, **kwargs)


def request_put(url_put, headers_put, data_put, **kwargs) -> requests:
    """
    http之put请求
    :param url_put:
    :param headers_put:
    :param data_put:
    :param kwargs:
    :return:
    """
    return requests.put(url=url_put, headers=headers_put, data=data_put, **kwargs)


def request_delete(url_de, headers_de, data_de, **kwargs) -> requests:
    """
    http之delete请求
    :param url_de: 请求url
    :param headers_de: 请求头
    :param data_de: 请求参数
    :param kwargs: 其他参数
    :return:
    """
    return requests.delete(url=url_de, headers=headers_de, data=data_de, **kwargs)

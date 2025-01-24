import time
from datetime import datetime


def get_current_timestamp() -> int:
    """
    获取当前时间戳（毫秒级）
    :return: 当前时间戳
    """
    return int(time.time() * 1000)


def format_timestamp_to_string(timestamp: int, format: str = "%Y-%m-%d %H:%M:%S") -> str:
    """
    将时间戳格式化为字符串
    :param timestamp: 时间戳
    :param format: 时间格式，默认为"%Y-%m-%d %H:%M:%S"
    :return: 格式化后的时间字符串
    """
    return time.strftime(format, time.localtime(timestamp / 1000))


def format_string_to_timestamp(time_str: str, format: str = "%Y-%m-%d %H:%M:%S") -> int:
    """
    将时间字符串转换为时间戳
    :param time_str: 时间字符串
    :param format: 时间格式，默认为"%Y-%m-%d %H:%M:%S"
    :return: 时间戳（毫秒级）
    """
    try:
        datetime_format = datetime.strptime(time_str, format)
        timestamp = int(datetime_format.timestamp() * 1000)
        return timestamp
    except ValueError as exc:
        raise ValueError('日期格式错误，需要传入的格式为 "%Y-%m-%d %H:%M:%S"') from exc


if __name__ == '__main__':
    current_timestamp = get_current_timestamp()
    print("当前时间戳:", current_timestamp)

    # 格式化时间戳为字符串，默认格式
    formatted_time = format_timestamp_to_string(current_timestamp)
    print("格式化时间:", formatted_time)

    # 格式化时间戳为字符串，自定义格式
    custom_format_time = format_timestamp_to_string(current_timestamp, "%Y-%m-%d %H:%M:%S %A")
    print("自定义格式化时间:", custom_format_time)

    # 将时间字符串转换为时间戳
    time_str = "2023-10-24 15:30:00"
    timestamp = format_string_to_timestamp(time_str)
    print("来自字符串的时间戳:", timestamp)

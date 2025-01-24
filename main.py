# 这是一个示例 Python 脚本。
import json
import socket

from apiconfig.api_base import ApiBase
from apiconfig.api_object import ApiObject

from utils.common_tool.common_data import find_value
from utils.excel_tool.excel_control import get_all_excel


# 按 ⌃R 执行或将其替换为您的代码。
# 按 双击 ⇧ 在所有地方搜索类、文件、工具窗口、操作和设置。


def print_hi(name):
    # 在下面的代码行中使用断点来调试脚本。
    print(f'Hi, {name}')  # 按 ⌘F8 切换断点。


# 按间距中的绿色按钮以运行脚本。
if __name__ == '__main__':
    qwe = {
        "code": 0,
        "message": "\u6210\u529f",
        "data": {
            "userId": 113415,
            "avatar": "https:\/\/yt3.ggpht.com\/ytc\/APkrFKYOJESfGdQafXf7OXaomDbd9JiCjG5_TxlfY38qLg=s88-c-k-c0x00ffffff-no-rj",
            "userName": "KoiFishu",
            "userEmail": "",
            "location": "1",
            "wechat": "55test1",
            "phone": "",
            "whatsApp": "",
            "line": "",
            "tags": [],
            "category": [],
            "channel": [
                {
                    "channelId": 1,
                    "channelName": "YouTube"
                }
            ]
        }
    }

    asd = find_value(qwe, 'data')
    print(asd)

# 访问 https://www.jetbrains.com/help/pycharm/ 获取 PyCharm 帮助

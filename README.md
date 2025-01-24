# API自动化测试框架介绍
## 前言

[框架同步钉钉文档](https://alidocs.dingtalk.com/i/nodes/m9bN7RYPWdlmbkzBf7RvlDqMWZd1wyK0?utm_scene=team_space)；

介于项目迭代更替，功能交互严重，回归测试环节中压力大时间长，不做回归测试不能确保证新老版本、新老功能完善且不冲突，估运用自动化实现新版本老功能的回归，缩短本次迭代测试范围，减少回归测试阶段耗时低效问题。需要掌握Python/pytest入门。

## 目标

本套自动化测试框架主要确保上线前重要API回归测试中，P0级别用例无BUG，P1级别用例BUG不超过5%。测试环境中项目主流程、主功能、无异常。

## 特点&优势&实现功能

- 特点
  - 采用ApiObject设计模式、运用Excel管理、快上手快熟练
- 优势
  - 数据/代码分离、API类可扩展、层级可扩展、灵活测试类
- 实现功能
  - 测试数据Excel驱动、单接口多场景断言、多接口依赖、更多可见Utils

## 技术栈&工具

- 代码
  - Python + PyTest + allure + yaml
- 工具
  - pycharm + Jenkins

## 框架目录结构

``` lua
API自动化测试框架目录结构示例：

|
|-- api_config               # 接口配置目录
|   └──  base_api.py         # api基类
|   └──  set_request.py      # 处理请求类
|
|-- common                # 公共目录
|   └──  config.yaml      # 配置文件
|   └──  setting.py       # 公共方法
|
|-- data                    # 测试数据存放目录，如Excel文件
|   └── object—name         # 项目名称
|       └── object.xslx     # 测试数据
|   └──  ...
|
|-- logs                         # 日志目录
|   └──  info.log                # info级别日志
|   └──  debug.log               # debug级别日志
|   └──  ...
|
|-- report                       # 测试报告本地目录
|   └── html                     # html格式报告
|   └── tmp                      # tmp格式报告
|   └── ...
|
|-- test_case                    # 测试类，层级可对比data目录
|   └──  object_name             # 项目名称
|       └──  test_xxx.py         # 测试py，可单个接口测试
|       └──  test_xxx.py         # 测试py，可多个接口测试(模块化)
|   └──  ...
|
|-- utils                         # 工具类目录
|   └── assert_tool               # 断言
|       └──  assert_control.py    # (待完善)
|   └──  excel_tool               # Excel
|       └──  excel_control.py     # 获取excel对象数据
|   └──  logging_tool             # 日志
|       └──  log_control.py       # 日志配置
|       └──  log_decorator.py     # api日志打印
|   └──  param_tool               # 参数
|       └──  global_params.py     # 全局参数
|       └──  random_params.py     # 随机参数
|   └──  time_tool                # 时间
|       └──  time_control.py      # 处理对应时间格式
|   └──  yaml_tool                # yaml
|       └──  yaml_control.py      # 获取yaml对象数据
|
|-- main.py                 # 调试入口
|-- pytest.ini              # pytest配置文件
|-- requirements.txt        # 依赖包
|-- run.py                  # 全局test运行入口

```

## 编写步骤

1. git拉下代码之后，安装requirements.txt依赖包
   终端运行 pip install -r requirements.txt
2. 按照data/demo.xlsx文件格式编写测试用例
   1. ***tips：请求data、请求headers、预期，三个列内容需要为json格式***
3. 按照test_case/..test_demo.py文件编写测试代码
4. 单接口、单模块运行步骤step2下方main方法
5. 可以运用pytest灵活执行所需要的测试快
6. run.py为全框架执行，带有test_开头全运行并生成allure报告

## 日志

所有日志目前统一存放在logs目录下

## 其它功能扩展

- 可提需求
- 敬请期待

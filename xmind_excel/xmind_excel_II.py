"""
一期的架构是从百度上扒来别人分享的基础上进行了修改，
在使用一段时间后，发现不太稳定，有存在一些bug且for循环太多，因此进行重构
"""
"""
xmind的原来编写格式不变，增加优化
1、预期结果取消判断空处理；
2、增加前置步骤（需要允许为空的状态）；
3、导入导出的文件存放在本地，非git目录下；


需要考虑的点：
1、用例步骤vs预期结果支持一对多
2、保留用例等级

"""
from xmindparser import xmind_to_dict
import openpyxl
import logging
import logging.handlers
from logging.handlers import TimedRotatingFileHandler
from datetime import datetime
import datetime
import timeit

class GetXmindExcel():
    """
    xmind转json
    """
    def __init__(self):
        pass

    def get_xmind_json(self):
        """
        将xmind转json格式
        :return:
        """
        pass

    def save_json_to_excel(self, xmind_json):
        """
        将json文件保持在本地文件
        :return:
        """
        pass
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
    def __init__(self,xmind_name, excel_name):

        self.xmind_name = xmind_name + ".xmind"
        self.excel_name = xmind_name + excel_name

        self.path = '/wuting'

        #计算测试用例的条数
        self.row_num = 1

    def xmind_title(self, value):
        """获取xmind标题内容"""
        return value['title']

    def numberLen(self, value):
        try:
            return len(value['topics'])
        except KeyError:
            return 0


    def get_xmind_json(self):
        """
        将xmind转json格式
        :return:
        """

        #获取到xmind的所有数据
        self.XmindContent = xmind_to_dict(self.xmind_name)[0]['topic']
        self.XmindTitle = self.xmind_title(self.XmindContent)
        FeaturesNUM = self.numberLen(self.XmindContent)
        TestName = ''  # 模块+标题名称
        # for i in range(FeaturesNUM):

    def save_json_to_excel(self, xmind_json):
        """
        将json文件保持在本地文件
        :return:
        """
        pass

if __name__ == '__main__':
    UserName = "不使用"#input('UserName:')
    XmindFile = input('XmindFile:') #xmind_excel.xmind

    if XmindFile == 'wt':
        XmindFile = 'xmind_excel'

    ExcelFile = input('ExcelFile:') #采销首页
    GetXmindExcel(XmindFile, ExcelFile).get_xmind_json()
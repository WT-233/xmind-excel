# -*- coding: utf-8 -*-
"""
一期的架构是从百度上扒来别人分享的基础上进行了修改，
在使用一段时间后，发现不太稳定，有存在一些bug且for循环太多，因此进行重构
"""
"""
xmind的原来编写格式不变，增加优化
1、预期结果取消判断空处理；
2、增加前置步骤（需要允许为空的状态）；
3、导入导出的文件存放在本地，非git目录下；
# -*- coding: utf-8 -*-
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

from xmind_excel.get_dic_value import GetDicValue

class GetXmindExcel(object):
    """
    xmind转json
    """
    def __init__(self,xmind_name, excel_name):

        self.xmind_name = xmind_name + ".xmind"
        self.excel_name = xmind_name + str(excel_name)

        self.path = '/wuting'

        #计算测试用例的条数
        # self.row_num = 1
        # self.xmind_file_dic = {}
        # self.xmind_file_name_dic = {}

        self.GET_DIC_VALUE = GetDicValue()

        self.xmind_dic = {}

        self.leves = {
            "leve0": "leve0",
            "leve1": "leve0",
            "leve2": "leve0",
            "leve3": "leve0",
            "leve4": "leve0"
        }

    def get_xmind_json(self):
        """
        将xmind转json格式
        :return:
        """

        #获取到xmind的基本信息
        xmind_file_dic = xmind_to_dict(self.xmind_name)[0]['topic']

        xmind_title = self.GET_DIC_VALUE.get_value_key_is_title(xmind_file_dic)
        xmind_topics = self.GET_DIC_VALUE.get_value_key_is_topics(xmind_file_dic)

        #用列表的形式存放模块信息
        moudle_names_list = []
        moudle_topics_list = []
        for i in xmind_topics:
            moudle_names_list.append(self.GET_DIC_VALUE.get_value_key_is_title(i))
            moudle_topics_list.append(self.GET_DIC_VALUE.get_value_key_is_topics(i))

        for j in range(0, len(moudle_names_list)):
            """
            获取模块名称
            """
            moudle_name = moudle_names_list[j]
            moudle_topics = moudle_topics_list[j]



            for t1 in range(0, len(moudle_topics)):
                """
                获取标题1、用例等级
                """
                title_names = ["","",""]
                title_leve = ""

                title_names[0] = self.GET_DIC_VALUE.get_value_key_is_title(moudle_topics[t1])
                title_topics = self.GET_DIC_VALUE.get_value_key_is_topics(moudle_topics[t1])
                if "makers" in title_topics[0].values():
                    title_leve = title_topics[0]["makers"]

                for t2 in title_topics:
                    """
                    获取标题2
                    """
                    title_names[1] = self.GET_DIC_VALUE.get_value_key_is_title(t2)
                    title_topics = self.GET_DIC_VALUE.get_value_key_is_topics(t2)

                    for t3 in title_topics:
                        """
                        获取标题3
                        """
                        title_names[2] = self.GET_DIC_VALUE.get_value_key_is_title(t3)
                        title_topics = self.GET_DIC_VALUE.get_value_key_is_topics(t3)

                        title_name = title_names[0] + title_names[1] + title_names[2]
                        title_preset = ""
                        title_step = ""
                        if title_topics is not None:
                            for s in range(0, len(title_topics)):
                                """
                                获取执行步骤
                                {'title': 'A1-1用例步骤1', 'topics': [{'title': 'A1-1预期结果1'}]}
                                """
                                step = ""

                                if len(title_topics) > 1:
                                    step = "步骤" + str(s+1) + ": "
                                    title_step += step + self.GET_DIC_VALUE.get_value_key_is_title(title_topics[s]) + "\n"
                                else:
                                    title_step += self.GET_DIC_VALUE.get_value_key_is_title(title_topics[s]) + "\n"
                                title_presets = self.GET_DIC_VALUE.get_value_key_is_topics(title_topics[s])

                                if title_presets is not None:
                                    """
                                    获取预期结果
                                    """

                                    for p in range(0, len(title_presets)):
                                        if len(title_presets) == 1 or p == 0:
                                            title_preset += step + self.GET_DIC_VALUE.get_value_key_is_title(
                                                title_presets[p])
                                            if (p == 0):
                                                title_preset += "\n"
                                        else:
                                            title_preset += self.GET_DIC_VALUE.get_value_key_is_title(title_presets[p]) + "\n"


                        xmind_dic = {}
                        xmind_dic["moudle"] = moudle_name
                        xmind_dic["title"] = title_name
                        xmind_dic["step"] = title_step
                        xmind_dic["preset"] = title_preset
                        print(xmind_dic)



























        # for i in range(FeaturesNUM):

    def save_json_to_excel(self, xmind_json):
        """
        将json文件保持在本地文件
        :return:
        """
        pass



if __name__ == '__main__':
    UserName = "不使用"#input('UserName:')
    XmindFile = 'wt' #input('XmindFile:') #xmind_excel.xmind

    if XmindFile == 'wt':
        XmindFile = 'xmind_excel'

    ExcelFile = "0"#input('ExcelFile:') #采销首页
    GetXmindExcel(XmindFile, ExcelFile).get_xmind_json()
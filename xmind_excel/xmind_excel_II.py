# -*- coding: utf-8 -*-
from xmindparser import xmind_to_dict
from xmind_excel.excel_write_read import ExcelWriteRead
from xmind_excel.get_dic_value import GetDicValue
"""
一期的架构是从百度上扒来别人分享的基础上进行了修改，
在使用一段时间后，发现不太稳定，有存在一些bug且for循环太多，因此进行重构
"""
"""
xmind的原来编写格式不变，增加优化
1、预期结果取消判断空处理；___已支持
2、导入导出的文件存放在本地，非git目录下；___已支持
3、用例步骤vs预期结果支持一对多___已支持
4、支持输出用例等级；___已支持
5、支持输出用例类型；___已支持
6、支持输出前置步骤；___已支持
"""

"""
20201223：
1、前置步骤可以在最后一个步骤识别
2、修正了用例等级 & 用例类型点的标识位置（标题1-3 都可以识别，标题3优先级大于标题1和2）

20201225
1、模块识别为需求id列
2、
"""



class GetXmindExcel(object):
    """
    xmind转json
    """
    def __init__(self,xmind_name, excel_name, is_need_num):

        self.xmind_name = xmind_name + ".xmind"
        self.excel_name = xmind_name + str(excel_name)

        self.path = '/Users/wuting/Documents/测试工作/测试用例/xmind_excel/'
        self.xmind_path = self.path + self.xmind_name

        # 模块名称是否为需求id
        self.is_need_num = is_need_num

        self.GET_DIC_VALUE = GetDicValue()

        self.rowNum = 1
        self.EXCEL_WRITE_READ = ExcelWriteRead(self.path, self.excel_name)


        self.xmind_list = []
        self.LEVEL_DIC = {
            "priority-1": "P0：确保系统基本功能及主要功能的测试",
            "priority-2": "P1：确保系统功能的完善方面的测试用例（包括一些异常数据、异常操作等）",
            "priority-3": "P2：界面UI方面的测试用例",
            "priority-4": "P3：确保系统兼容性方面的测试用例（如与手机本身功能的兼容等）",
            "priority-5": "P4：用户体验方面的测试用例"
        }
        # 暂不投入使用
        self.TAG_DIC = {
            "0": "功能测试",
            "1": "接口用例",
            "2": "性能用例",
            "3": "安全性用例",
            "4": "兼容性用例",
            "5": "交互（UI/UE）用例",
            "6": "配置用例",
            "7": "组件用例",
            "8": "文档用例",
            "9": "适配用例",
            "10": "冒烟用例",
        }

        self.row_num = 1

    def get_level_word(self, leve):
        """
        放测试用例等级的具体文字描述
        :param leve:
        :return:
        """

        if leve in self.LEVEL_DIC.keys():
            return self.LEVEL_DIC[leve]
        else:
            return ""

    def get_xmind_json(self):
        """
        将xmind转json格式
        :return:
        """

        # 获取到xmind的基本信息
        xmind_file_dic = xmind_to_dict(self.xmind_path)[0]['topic']

        xmind_title = self.GET_DIC_VALUE.get_value_key_is_title(xmind_file_dic)
        xmind_topics = self.GET_DIC_VALUE.get_value_key_is_topics(xmind_file_dic)

        # 用列表的形式存放模块信息
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

            self.rowNum = 1
            for t1 in range(0, len(moudle_topics)):
                """
                获取用例等级、标题1
                """
                title_level = ''
                title_tag = ''
                # 用例等级
                if "makers" in moudle_topics[t1].keys():
                    for t1_i in moudle_topics[t1]["makers"]:
                        if len(t1_i) > 7:
                            if t1_i[:8] == 'priority':
                                title_level = t1_i
                                title_level = self.get_level_word(title_level)
                # 用例类型
                if "labels" in moudle_topics[t1].keys():
                    title_tag = moudle_topics[t1]["labels"][0]


                title_names = ["", "", ""]

                title_names[0] = self.GET_DIC_VALUE.get_value_key_is_title(moudle_topics[t1])
                title_topics1 = self.GET_DIC_VALUE.get_value_key_is_topics(moudle_topics[t1])

                for t2 in title_topics1:
                    """
                    获取标题2
                    """
                    # 用例等级
                    if "makers" in title_topics1[t2].keys():
                        for t2_1 in title_topics1[t2]["makers"]:
                            if len(t2_1) > 7:
                                if t2_1[:8] == 'priority':
                                    title_level = t2_1
                                    title_level = self.get_level_word(title_level)
                    # 用例类型
                    if "labels" in title_topics1[t2].keys():
                        title_tag = title_topics1[t2]["labels"][0]

                    title_names[1] = self.GET_DIC_VALUE.get_value_key_is_title(t2)
                    title_topics2 = self.GET_DIC_VALUE.get_value_key_is_topics(t2)

                    for t3 in range(0, len(title_topics2)):
                        """
                        用例等级、用例类型
                        获取标题3
                        """

                        # 用例等级
                        if "makers" in title_topics2[t3].keys():
                            for t3_1 in title_topics2[t3]["makers"]:
                                if len(t3_1) > 7:
                                    if t3_1[:8] == 'priority':
                                        title_level = t3_1
                                        title_level = self.get_level_word(title_level)
                        # 用例类型
                        if "labels" in title_topics2[t3].keys():
                            title_tag = title_topics2[t3]["labels"][0]

                        title_names[2] = self.GET_DIC_VALUE.get_value_key_is_title(title_topics2[t3])
                        title_topics3 = self.GET_DIC_VALUE.get_value_key_is_topics(title_topics2[t3])

                        # title_name = title_names[0] + " - " + title_names[1] + " - " + title_names[2]
                        title_name = ''
                        for t in range(0, len(title_names)):
                            if t == 0:
                                title_name += title_names[t]
                            else:
                                title_name += ' - ' + title_names[t]

                        title_preset = ""
                        title_step = ""
                        title_pre_step = ""
                        if title_topics3 is not None:
                            step_num = 1
                            for s in range(0, len(title_topics3)):
                                """
                                获取前置步骤、执行步骤
                                """
                                # 前置步骤在最前
                                if s == 0 and "makers" in title_topics3[s].keys():
                                    if 'flag-red' in self.GET_DIC_VALUE.get_value_key_is_makers(title_topics3[s]):
                                        title_pre_step = self.GET_DIC_VALUE.get_value_key_is_title(title_topics3[s])
                                        continue
                                # 前置步骤在最末
                                if s == len(title_topics3)-1 and "makers" in title_topics3[s].keys():
                                    if 'flag-red' in self.GET_DIC_VALUE.get_value_key_is_makers(title_topics3[s]):
                                        title_pre_step = self.GET_DIC_VALUE.get_value_key_is_title(title_topics3[s])
                                        continue
                                step = ""
                                if len(title_topics3) > 1:
                                    step = "步骤" + str(step_num) + ": "
                                    step_num += 1
                                    title_step += step + self.GET_DIC_VALUE.get_value_key_is_title(title_topics3[s]) + "\n"
                                else:
                                    title_step += self.GET_DIC_VALUE.get_value_key_is_title(title_topics3[s]) + "\n"
                                title_presets = self.GET_DIC_VALUE.get_value_key_is_topics(title_topics3[s])

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
                        xmind_dic["用例目录"] = ""
                        xmind_dic["用例名称"] = title_name
                        xmind_dic["需求ID"] = ""
                        xmind_dic["前置条件"] = title_pre_step
                        xmind_dic["用例步骤"] = title_step
                        xmind_dic["预期结果"] = title_preset
                        xmind_dic["用例类型"] = title_tag
                        xmind_dic["用例状态"] = ""
                        xmind_dic["用例等级"] = title_level
                        xmind_dic["创建人"] = ""
                        xmind_dic["用例类型自定义"] = ""

                        if self.is_need_num == 1:
                            # if type(moudle_name) is int:
                            xmind_dic["需求ID"] = moudle_name
                            # else:
                            #     print("需求id非int格式，需求ID生成失败")
                            #     return "需求id非int格式，需求ID生成失败"

                        # xmind_dic["模块"] = moudle_name

                        self.xmind_list.append(xmind_dic)
                        # print(xmind_dic["用例等级"], xmind_dic["用例名称"],xmind_dic["用例类型"])

                        self.rowNum = self.rowNum + 1
                        self.EXCEL_WRITE_READ.excel_write(self.rowNum, xmind_dic, moudle_name)


if __name__ == '__main__':

    UserName = "不使用"
    # UserName = input('UserName:')

    XmindFile = input('XmindFile:')
    if XmindFile == 'wt' or XmindFile == 'WT':
        XmindFile = 'es索引优化-导出'

    ExcelFile = input('ExcelFile:')

    IsNeed = input('模块是否为需求id?  [1：是，需求id格式如：1041169]： ')
    if IsNeed == 1 or '1':
        GetXmindExcel(XmindFile, ExcelFile, 1).get_xmind_json()
    else:
        GetXmindExcel(XmindFile, ExcelFile, 0).get_xmind_json()
# -*- coding: utf-8 -*-
"""
获取dic的value
"""
class GetDicValue(object):

    def __init__(self, error_text = None,error_value = ''):
        self.error_text = error_text
        self.error_value = error_value

    def get_value_key_is_title(self, i_dic, error_text = None):
        """
        获取dic，key=title,的value
        :param value:
        :return:
        """
        if error_text is None:
            error_text = self.error_value
        try:
            if i_dic['title'] is None:
                return error_text

            return i_dic['title']
        except Exception as e:
            return error_text

    def get_value_key_is_topic(self, i_dic, error_text = None):
        """
        获取dic，key=topic,的value
        :param value:
        :return:
        """
        if error_text is None:
            error_text = self.error_value
        try:
            return i_dic['topic']
        except Exception as e:
            return error_text

    def get_value_key_is_topics(self, i_dic, error_text = None):
        """
        获取dic，key=topics,的value
        :param value:
        :return:
        """
        if error_text is None:
            error_text = self.error_value
        try:
            return i_dic['topics']
        except Exception as e:
            return error_text

    def get_value_key_is_makers(self, i_dic, error_text = None):
        """
        获取dic，key=topics,的value
        :param value:
        :return:
        """
        if error_text is None:
            error_text = self.error_value
        try:
            return i_dic['makers']
        except Exception as e:
            return error_text

if __name__ == '__main__':
    pass
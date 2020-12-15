# -*- coding: utf-8 -*-
"""
获取dic的value
"""
class GetDicValue(object):

    def __init__(self, error_text = None):
        self.error_text = error_text

    def get_value_key_is_title(self, i_dic):
        """
        获取dic，key=title,的value
        :param value:
        :return:
        """

        try:
            return i_dic['title']
        except Exception as e:
            return self.error_text

    def get_value_key_is_topic(self, i_dic):
        """
        获取dic，key=topic,的value
        :param value:
        :return:
        """
        try:
            return i_dic['topic']
        except Exception as e:
            return self.error_text

    def get_value_key_is_topics(self, i_dic):
        """
        获取dic，key=topics,的value
        :param value:
        :return:
        """
        try:
            return i_dic['topics']
        except Exception as e:
            return self.error_text
if __name__ == '__main__':
    pass
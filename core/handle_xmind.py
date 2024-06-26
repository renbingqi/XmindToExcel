# -*- ecoding: utf-8 -*-
# @ModuleName: handel_xmind
# @Author: Rex
# @Time: 2022/12/8 6:54 下午
import xmindparser


class HandleXmind():
    def __init__(self, xmind_file):
        self.xmind_file = xmind_file
        self.sheetName = None
        self.case_list = []
        self.maxModule = 0

    def handle_xmind(self):
        """
        通过xmindparser库将xmind中的数据转换成dict类型，拿到第一个节后后的所有数据
        [{'title': '一级模块', 'makers': ['priority-1'], 'topics': [{'title': '二级模块', 'makers': ['priority-2'],
        'topics': [{'title': '标题：demo-一级模块测试\n前置：前置可有可无\n步骤：进入一级模块\n预期：正常进入一级模块', 'makers': ['task-done'],
        'topics': [{'title': '标题：demo-二级模块测试-01\n前置：前置可有可无\n步骤：进入二级模块\n预期：正常进入二级模块', 'note': '进入二级模块失败',
        'makers': ['symbol-attention']}, {'title': '标题：demo-二级模块测试-02\n前置：前置可有可无\n步骤：进入二级模块\n预期：正常进入二级模块',
        'makers': ['symbol-attention'], 'labels': ['进入二级模块失败']}]}]}]}, {'title': '一级模块', 'makers': ['priority-1'],
        'topics': [{'title': '二级模块', 'makers': ['priority-2'], 'topics': [{'title': '三级模块', 'makers': ['priority-3'],
        'topics': [{'title': '标题：这是一个没有前置步骤的测试用例\n步骤：进入一级模块\n预期：正常进入一级模块', 'makers': ['task-done']},
         {'title': '没有标记 & 没有标题（不是Case）的节点将会被过滤，可以再这里写点逻辑性的内容，整理思路',
         'topics': [{'title': '标题：这里才是真正的Case\n步骤：进入三级模块\n预期：成功', 'makers': ['task-done']}]}]}]}]}]
        :return:
        """
        dict_data = xmindparser.xmind_to_dict(self.xmind_file)
        all_data = dict_data[0]['topic']
        # xmind内容主题，可用该名字作为最后Excel报告的文件名
        self.sheetName = all_data['title']
        # 获取所有的1级节点
        topics = all_data['topics']
        self.get_all_topic_data(topics, {})

    def get_all_topic_data(self, data, dic):
        """
        拿到所有topic的数据，并传递有效的case值 type dict
        :param data: 需要处理的数据
        :param dic: 有效的case type:dict
        :return:
        """
        # print(data)
        dict_case = dic
        # 说明没有向下的分节点了
        if len(data) == 1:
            dict_data = data[0]
            # 包含标签说明是我们想要的节点,并从该节点中提取title的值
            self.get_title_data(dict_case, dict_data)
        # 说明有向下的节点
        else:
            for i in range(len(data)):
                if i == 0:
                    # 为了不改变第一次传进来的数据所以先copy一份用copy的数据进行组装
                    self.get_title_data(dic.copy(), data[i])
                else:
                    # 如果第一级节点包含了标签说明是我们想要的节点
                    if list(data[i].keys()).__contains__("makers"):
                        # 如果该节点包含 1 标记说明新的1级模块，此时不需要复制之前的case
                        if data[i]['makers'] == ['priority-1']:
                            new_dict_case = {}
                            self.get_title_data(new_dict_case, data[i])
                        # 如果该节点包含标签但不是 1，比如说 2，3，4 说明之前已经有出现过1级标签了，此时需要复制之前的case
                        else:
                            new_dict_case = dict_case.copy()
                            self.get_title_data(new_dict_case, data[i])
                    # 如果是case节点也是我们想要的
                    elif "标题" in data[i]['title']:
                        new_dict_case = dict_case.copy()
                        self.get_title_data(new_dict_case, data[i])
                    # 如果该节点没有标签 且 标题也不在其中（不是case节点），则不是我们想要的节点，检查下是否有下级节点
                    else:
                        # 说明还有下一级节点,就直接递归
                        if list(data[i].keys()).__contains__("topics"):
                            new_dict_case = dict_case.copy()
                            self.get_all_topic_data(data[i]['topics'], new_dict_case)
                        # 如果没有下一级节点了，说明已经是最后的节点了，则把数据放到列表中

    def get_title_data(self, dict_case, dict_data):
        """
        处理数据，并拿到所有需要数据的titile值，并将数据填入有效的case值中（dict类型）（title对应的是用户在xmind中输入的数据）
        :param dict_case: 有效的case type dict
        :param dict_data: 需要处理的数据
        :return:
        """
        if list(dict_case.keys()).__contains__("title"):
            # 先复制之前的case
            new_dict_case = dict_case.copy()

            if list(dict_data.keys()).__contains__("makers"):
                if dict_data['makers'] == ['priority-1']:
                    new_dict_case['module-1'] = dict_data['title']
                    self.check_max_module(1)
                elif dict_data['makers'] == ['priority-2']:
                    new_dict_case['module-2'] = dict_data['title']
                    self.check_max_module(2)
                elif dict_data['makers'] == ['priority-3']:
                    new_dict_case['module-3'] = dict_data['title']
                    self.check_max_module(3)
                elif dict_data['makers'] == ['priority-4']:
                    new_dict_case['module-4'] = dict_data['title']
                    self.check_max_module(4)

            # if "标题" in dict_data['title'] and "预期" in dict_data['title']:
            if "标题" in dict_data['title'] and "期望"or"预期" in dict_data['title']:
                case = dict_data['title']
                self.case_format(new_dict_case, case)
                # 设置case的状态
                self.set_case_status(new_dict_case, dict_data)
                # 提取case的note状态
                self.get_case_note_labels(new_dict_case, dict_data)
                self.case_list.append(new_dict_case)

            # 说明还有下一级节点,就直接递归
            if list(dict_data.keys()).__contains__("topics"):
                self.get_all_topic_data(dict_data['topics'], new_dict_case)
        else:
            if list(dict_data.keys()).__contains__("makers"):
                if dict_data['makers'] == ['priority-1']:
                    dict_case['module-1'] = dict_data['title']
                    self.check_max_module(1)
                elif dict_data['makers'] == ['priority-2']:
                    dict_case['module-2'] = dict_data['title']
                    self.check_max_module(2)
                elif dict_data['makers'] == ['priority-3']:
                    dict_case['module-3'] = dict_data['title']
                    self.check_max_module(3)
                elif dict_data['makers'] == ['priority-4']:
                    dict_case['module-4'] = dict_data['title']
                    self.check_max_module(4)

            # 如果发现topics里面的内容只有一级包含标题的内容时 则需要检查case前面是否已经有了
            # 当出现标题在节点中时，该节点也不一定是最后一个节点，也许是case后面还有case
            if "标题" in dict_data['title'] and "期望" or"预期" in dict_data['title'] :
            # if "标题" in dict_data['title'] and "预期" in dict_data['title']:
                # 如果case的字典里已经有case的字段了，说明这个是case后面的case
                if list(dict_case.keys()).__contains__("title"):
                    # 先复制之前的case
                    new_dict_case = dict_case.copy()
                    # 再把case字段的值替换掉
                    case = dict_data['title']
                    self.case_format(new_dict_case, case)
                    # 设置case的状态
                    self.set_case_status(new_dict_case, dict_data)
                    # 提取case的note状态
                    self.get_case_note_labels(new_dict_case, dict_data)
                    self.case_list.append(new_dict_case)
                else:
                    case = dict_data['title']
                    self.case_format(dict_case, case)
                    # 检查这个case中是否有makers标记并设置case的状态
                    self.set_case_status(dict_case, dict_data)
                    # 检查这个case中是否有note标记 有的话把note提取出来
                    self.get_case_note_labels(dict_case, dict_data)
                    # 先把前面的case加入到case列表中
                    self.case_list.append(dict_case)
            # 说明还有下一级节点,就直接递归
            if list(dict_data.keys()).__contains__("topics"):
                self.get_all_topic_data(dict_data['topics'], dict_case)
            # 如果没有下一级节点了，说明已经是最后的节点了，则把数据放到列表中

    def set_case_status(self, dict_case, dict_data):
        """
        解析xmind中实际的case数据，并根据实际情况赋予case状态值
        :param dict_case:
        :param dict_data:
        :return:
        """
        if list(dict_data.keys()).__contains__("makers"):
            maker = dict_data['makers']
            if maker.__contains__("task-done"):
                dict_case['case_status'] = "PASS"
                # 兼容mac&windows版本
            elif maker.__contains__("symbol-attention") or maker.__contains__("symbol-exclam"):
                dict_case['case_status'] = "FAIL"

            else:
                dict_case['case_status'] = "N/A"
            # 增加用例优先级的判断，作用是否要用于回归
            if maker.__contains__("priority-1"):
                dict_case['regression'] = "1"
            elif maker.__contains__("priority-2"):
                dict_case['regression'] = "2"
            elif maker.__contains__("priority-3"):
                dict_case['regression'] = "3"
            else:
                dict_case['regression'] = "N/A"

        else:
            dict_case['case_status'] = "N/A"

    def get_case_note_labels(self, dict_case, dict_data):
        """
        处理case中的note和labels数据，一般该数据是指case失败后的备注
        :param dict_case:
        :param dict_data:
        :return:
        """
        if list(dict_data.keys()).__contains__("note"):
            note = dict_data['note']
            dict_case['note'] = note
        elif list(dict_data.keys()).__contains__("labels"):
            labels = dict_data['labels']
            str_labels = ",".join(labels)
            dict_case['note'] = str_labels
        else:
            dict_case['note'] = ""

    def check_max_module(self, module):
        """
        对最大模块数进行更新
        :param module:
        :return:
        """
        if module > self.maxModule:
            self.maxModule = module

    def case_format(self, dict_case, case):
        """
        处理case，将标题，前置，步骤，预期等解析出来并添加到case中
        :param dict_case:
        :param case:
        :return:
        """
        if "：" in case or ":" in case:
            replace_case = case.replace("：", ":")
            if "前置:" in replace_case:
                indexPreconditions = replace_case.index("前置:")
                indexTestStep = replace_case.index("步骤:")

                try:
                    indexExpected_Result = replace_case.index("期望:")
                except:
                    indexExpected_Result = replace_case.index("预期:")
                # indexExpected_Result = replace_case.index("预期:")
                title = replace_case[3:indexPreconditions]
                Preconditions = replace_case[indexPreconditions + 3:indexTestStep]
                TestStep = replace_case[indexTestStep + 3:indexExpected_Result]
                ExpectedResult = replace_case[indexExpected_Result + 3:]
                dict_case["title"] = title.rstrip()
                dict_case["preconditions"] = Preconditions.rstrip()
                dict_case["TestStep"] = TestStep.rstrip()
                dict_case["ExpectedResult"] = ExpectedResult.rstrip()
            else:
                indexTestStep = replace_case.index("步骤:")
                try:
                    indexExpected_Result = replace_case.index("期望:")
                except:
                    indexExpected_Result = replace_case.index("预期:")
                title = replace_case[3:indexTestStep]
                TestStep = replace_case[indexTestStep + 3:indexExpected_Result]
                ExpectedResult = replace_case[indexExpected_Result + 3:]
                dict_case["title"] = title.rstrip()
                dict_case["preconditions"] = ''
                dict_case["TestStep"] = TestStep.rstrip()
                dict_case["ExpectedResult"] = ExpectedResult.rstrip()


if __name__ == '__main__':
    xmind_file = "../Demo.xmind"
    xmind_handler = HandleXmind(xmind_file)
    xmind_handler.handle_xmind()
    print(xmind_handler.case_list)

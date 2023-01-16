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
        dict_data = xmindparser.xmind_to_dict(self.xmind_file)
        all_data = dict_data[0]['topic']
        # xmind内容主题，可用该名字作为最后Excel报告的文件名
        self.sheetName = all_data['title']
        # 获取所有的1级节点
        topics = all_data['topics']
        self.get_all_topic_data(topics, {})

    def get_all_topic_data(self, data, dic):
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
                    self.get_title_data(dict_case, data[i])
                else:
                    new_dict_case = {}
                    # 如果第一级节点包含了标签说明是我们想要的节点
                    if list(data[i].keys()).__contains__("makers"):
                        # 如果该节点包含 1 标记说明新的1级模块，此时不需要复制之前的case
                        if data[i]['makers'] == ['priority-1']:
                            pass
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
                            self.get_all_topic_data(data[i]['topics'], new_dict_case)
                        # 如果没有下一级节点了，说明已经是最后的节点了，则把数据放到列表中

    def get_title_data(self, dict_case, dict_data):
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
        if "标题" and "预期" in dict_data['title']:
            # 如果case的字典里已经有case的字段了，说明这个是case后面的case
            if list(dict_case.keys()).__contains__("title"):
                # 先复制之前的case
                new_dict_case = dict_case.copy()
                # 再把case字段的值替换掉
                case = dict_data['title']
                self.case_format(new_dict_case,case)
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
                self.get_case_note_labels(dict_case,dict_data)
                # 先把前面的case加入到case列表中
                self.case_list.append(dict_case)
        # 说明还有下一级节点,就直接递归
        if list(dict_data.keys()).__contains__("topics"):
            self.get_all_topic_data(dict_data['topics'], dict_case)
        # 如果没有下一级节点了，说明已经是最后的节点了，则把数据放到列表中

    def set_case_status(self, dict_case, dict_data):
        if list(dict_data.keys()).__contains__("makers"):
            maker = dict_data['makers']
            if maker == ['task-done']:
                dict_case['case_status'] = "PASS"
            elif maker == ['symbol-attention']:
                dict_case['case_status'] = "FAIL"
            else:
                dict_case['case_status'] = "N/A"
        else:
            dict_case['case_status'] = "N/A"

    def get_case_note_labels(self, dict_case, dict_data):
        if list(dict_data.keys()).__contains__("note") :
            note = dict_data['note']
            dict_case['note'] = note
        elif list(dict_data.keys()).__contains__("labels") :
            labels = dict_data['labels']
            str_labels = ",".join(labels)
            dict_case['note'] = str_labels
        else:
            dict_case['note'] = ""

    def check_max_module(self, module):
        if module > self.maxModule:
            self.maxModule = module

    def case_format(self,dict_case,case):
        if "前置" in case:
            indexPreconditions=case.index("前置")
            indexTestStep = case.index("步骤")
            indexExpected_Result = case.index("预期")
            title = case[3:indexPreconditions]
            Preconditions = case[indexPreconditions + 3:indexTestStep]
            TestStep = case[indexTestStep + 3:indexExpected_Result]
            ExpectedResult = case[indexExpected_Result + 3:]
            dict_case["title"] = title.rstrip()
            dict_case["preconditions"] = Preconditions.rstrip()
            dict_case["TestStep"] = TestStep.rstrip()
            dict_case["ExpectedResult"] = ExpectedResult.rstrip()
        else:
            indexTestStep = case.index("步骤")
            indexExpected_Result = case.index("预期")
            title = case[3:indexTestStep]
            TestStep = case[indexTestStep + 3:indexExpected_Result]
            ExpectedResult = case[indexExpected_Result + 3:]
            dict_case["title"] = title.rstrip()
            dict_case["preconditions"] = ''
            dict_case["TestStep"] = TestStep.rstrip()
            dict_case["ExpectedResult"] = ExpectedResult.rstrip()






if __name__ == '__main__':
    xmind_file = "Server SDK.xmind"
    xmind_handler = HandleXmind(xmind_file)
    xmind_handler.handle_xmind()
    for item in xmind_handler.case_list:
        print(item)


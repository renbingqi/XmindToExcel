# -*- ecoding: utf-8 -*-
# @ModuleName: handel_excel
# @Author: Rex
# @Time: 2023/1/15 8:15 下午
import xlwt
class HandleExcel():
    def __init__(self,fileName,filePath):
        # fileName既作为文件名也作为sheet name
        self.fileName=fileName
        self.workbook=xlwt.Workbook(encoding='utr-8')
        self.worksheet=self.workbook.add_sheet(fileName)
        self.case_demo={}
        self.filePath=filePath

    def generate_title(self,maxModule):
        title_list=["Number"]
        # 根据模块的级数生成对应的标题
        for item in range(maxModule):
            module = "Module-"+str(item + 1)
            title_list.append(module)
        title_list+=["Test Item","Preconditions","Test Step","Expected Result","Result","Note"]
        # 将生成的标题写入到excel中
        for i in range(len(title_list)):
            # 标题在第一行，所以行号都固定为0，列号对应标题列表的索引值
            self.worksheet.write(0, i, title_list[i])
            # 拿到所有标题的列号，并生成一个字典，存储每个标题及对应的列号
            if title_list[i] == "Test Item":
                self.case_demo ['title'] = i
            elif title_list[i] == "Test Step":
                self.case_demo['TestStep']=i
            elif title_list[i] == "Expected Result":
                self.case_demo['ExpectedResult']=i
            elif title_list[i] == "Result":
                self.case_demo['case_status']=i
            elif title_list[i] == "Note":
                self.case_demo['note']=i
            else:
                self.case_demo[title_list[i].lower()] = i
        print()
        self.workbook.save(f"{self.filePath}.xls")
    def write_data(self,data_list):
        try:
            for item in range(len(data_list)):
                case_id = item + 1
                self.worksheet.write(case_id,0 ,case_id)
                # 拿到所有case的key
                key_list=list(data_list[item].keys())
                for key in range(len(key_list)):
                    # 根据case的key 从 case_demo中拿到对应标题列号
                    # case的每个key都对应case_demo中的key
                    clo= self.case_demo[key_list[key]]
                    self.worksheet.write(case_id,clo ,data_list[item][key_list[key]])
            self.workbook.save(f"{self.filePath}.xls")
            return True
        except:
            return False

if __name__ == '__main__':
    excelHandle = HandleExcel("test")
    excelHandle.generate_title(3)
    excelHandle.write_data()





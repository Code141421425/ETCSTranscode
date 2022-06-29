import os
import xlrd


class Transcode:
    # class的基础语法模板
    classTemplate = ""
    # 每条数据的语法模板，用于反复添加数据，最后添加到classTemplate中
    singleDataTemplate = ""

    transcodePos = ""
    rowDataPos = ""

    PROJECT_PATH = os.path.dirname(__file__)
    DEFAULT_ROW_DATA_POS = PROJECT_PATH + "/ExecuteZone/RowData/"
    DEFAULT_TRANSCODE_POS = PROJECT_PATH + "/ExecuteZone/TranscodePos/"

    def __init__(self, rowDataPos=DEFAULT_ROW_DATA_POS, transcodePos=DEFAULT_TRANSCODE_POS):
        self.transcodePos = transcodePos
        self.rowDataPos = rowDataPos

    def run(self):
        pass


class ToCSharp(Transcode):
    def __init__(self, *args, **kwargs):
        super(ToCSharp, self).__init__(*args, **kwargs)

        self.fileNameList = []

        self.classTemplate = """
using System.Collections.Generic;
namespace data{
public static class %s{
        %s     
}} 
"""
        self.singleDataTemplate = """
    public static Dictionary<string,dynamic> %s = new Dictionary<string,dynamic>(){
%s
    };
"""
        self.getDataTemplate = """
using UnityEngine;
namespace data{
public class GetData:MonoBehaviour
{
%s
}}     
"""
        self.dataListTemplate = """
    public static Dictionary<int,Dictionary<string,dynamic> > %s = new Dictionary<int,Dictionary<string,dynamic>>(){
%s
    };
"""
        self.getDataParamTemplate = "    public %s %s = new %s();\n"
        self.paramTemplate = "        {\"%s\",%s},\n"
        self.dictData = ""
        self.listData = ""
        self.excel = None

    def __processByType(self, i, j):
        sheet = self.excel.sheet_by_index(0)
        rowData = sheet.cell_value(i, j)
        dataType = sheet.cell_value(1, j)

        # 根据excel中所填写的类型，处理生数据的格式
        if dataType == "string":
            result = "\"%s\"" % rowData

        elif dataType == "float":
            result = rowData+"f"

        elif dataType == "strings":
            rowData = ""+rowData
            listData = ""
            for data in rowData.split(","):
                listData += "\"%s\"," % data

            result = "new List<string>{%s}" % listData

        elif dataType == "ints":
            listData = ""
            for data in rowData.split(","):
                listData += "%s," % data
            result = "new List<int>{%s}" % listData

        elif dataType == "int-string":
            sheetT = self.excel.sheet_by_name("translate")
            paramType = sheet.cell_value(0, j)

            # 在翻译表第1行内，进行循环，得出
            typeIndex = -1
            for i in range(0, sheetT.row_len(0)):
                if sheetT.cell_value(0, i) == paramType:
                    typeIndex = i
                    break
            if typeIndex == -1:
                pass
            else:
                result = "\"%s\"" % sheetT.cell_value(int(rowData), typeIndex)

        else:
            result = rowData

        return result

    def transcodeOnce(self, fileName):
        # 用于单次从rowData到c#的转码，重复调用
        # print(self.rowDataPos + fileName)
        self.excel = xlrd.open_workbook_xls(self.rowDataPos + fileName)
        sheet = self.excel.sheet_by_index(0)

        for i in range(2, sheet.nrows):
            # 组装一行excel组成的数据
            paramData = ""
            for j in range(sheet.ncols):
                # 组成新的一个字典的一条数据
                if sheet.cell_value(0, j) != "Note":
                    paramData += self.paramTemplate % (sheet.cell_value(0, j), self.__processByType(i, j))

            # 通过数据和模板，组成一个新的字典，字典名称用excel第2列为字典名
            self.dictData += self.singleDataTemplate % (sheet.cell_value(i, 1), paramData)
            # 将id和Name做为键值对，加入listDict，用以生成最后的list
            self.listData += "        {%s,%s},\n" % (sheet.cell_value(i, 0), sheet.cell_value(i, 1))

    def transcode(self):
        try:
            files = os.listdir(self.rowDataPos)
        except FileNotFoundError:
            print("rowDataPos Error")
            return

        for file in files:
            # 确认后缀
            # if file[-4:] == "xlsx":
            #     self.transcodeOnce(file)
            if file[-3:] == "xls":
                fileName = file[:-4]
                self.fileNameList.append(fileName)
                self.transcodeOnce(file)
                # 增加classList的数据
                self.dictData += self.dataListTemplate % (fileName.lower()+"List", self.listData)
                completeData = self.classTemplate % (fileName, self.dictData)

                self.writeTranscode(fileName, completeData)
                self.dictData = ""
                self.listData = ""

    def writeTranscode(self, fileName, data):
        file = open(self.transcodePos + fileName + ".cs", mode="w")
        file.write(data)
        file.close()

    # def generateGetData(self):
    #     # 由于生成最后的getData.cs
    #     # 【已弃用】
    #     classData = ""
    #
    #     for fileName in self.fileNameList:
    #         fileName = ""+fileName
    #         classData += self.getDataParamTemplate % (fileName, fileName.lower(), fileName)
    #
    #     result = self.getDataTemplate % classData
    #
    #     return result

    def run(self):
        self.transcode()
        # self.writeTranscode("GetData", self.generateGetData())


if __name__ == '__main__':
    # ToCSharp(transcodePos=r"H:\WorkSpace\Unity\NG\NG1\Assets\Script\0_TechTest\Data/").run()
    ToCSharp(transcodePos=r"H:\WorkSpace\Unity\NG\NG1_v1\Assets\_Scripts\data/").run()
    print("excel To c# finish")


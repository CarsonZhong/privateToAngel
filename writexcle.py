#!/usr/bin/python3
#coding=gbk
import copy
import xlrd
import xlwt
from common.readconfig import ReadConfig

# ���ݵ���С��Ԫ��ʽ
class DataInfo:
    """�ձ����ݽṹ"""
    class Struct(object):
        def __init__(self, sClo = -1, dClo = -1, dataStyle = 'S', SData = ''):
            self.sClo = sClo            # ����Դ��
            self.dClo = dClo            # ����Ŀ����
            self.dataStyle = dataStyle  # ���ݸ�ʽ
            self.SData = SData          # �ַ�������

        def make_struct(self, sClo, dClo, dataStyle, SData):
            return self.Struct(sClo, dClo, dataStyle, SData)
# �����ݸ�ʽ
class Rowdata(object):
    """ԭʼ���ݱ�"""
    class Struct(object):
        def __init__(self, rownum = -1, key = ''):
            self.dataList = []
            self.rownum = rownum
            self.key = key

    def make_struct(self, rownum = -1, key = ''):
        return self.Struct(rownum, key)

class mapdata:
    """ӳ���"""
    name = ''
    id = ''


# �������ݸ�ʽ
class alldata:
    """�������ݱ�"""
    def __init__(self):
        self.dataList = []

    def AddOneData(self, data):
        self.dataList.append(data)

def getMapIdByDict(configdata):
    mapxlsname = configdata.get_map("mapxlsname")
    RowStart = int(configdata.get_map("RowStart"))
    nameclos = int(configdata.get_map("nameclos"))
    mapidclos = int(configdata.get_map("mapidclos"))

    srcbook = xlrd.open_workbook(mapxlsname)
    srcSheet = srcbook.sheets()[0]
    sumrows = srcSheet.nrows

    dict = {}
    for row in range(RowStart, sumrows):
        name = srcSheet.cell(row, nameclos).value
        id = srcSheet.cell(row, mapidclos).value
        dict[name] = id

    return dict

def getStyle(configdata):
    stylexlsname = configdata.get_style("stylexlsname")
    dataRow = int(configdata.get_style("dataRow"))
    ColStart = int(configdata.get_style("ColStart"))

    styleRow = dataRow + 1
    stylebook = xlrd.open_workbook(stylexlsname)
    styleSheet = stylebook.sheets()[0]

    styleCol = ColStart

    styleList = Rowdata().make_struct()
    while styleSheet.cell(dataRow, styleCol).value:
        cell = styleSheet.cell(dataRow, styleCol)
        data = DataInfo().Struct()

        if cell.value != 'MAP' :
            data.sClo = ord(cell.value) - ord('A')

        data.dClo = styleCol
        data.dataStyle = styleSheet.cell(styleRow, styleCol).value

        styleList.dataList.append((data))

        styleCol = styleCol + 1

    return styleList

def getName(dataToList):
    for data in dataToList:
        if data.dataStyle == 'N':
            return data.SData
    return ''

def getKey(dataToList, data):
    if data.dataStyle == 'K':
        return data.SData
    return dataToList.key

def combineSame(TargetData, dataToList):
    hasSame = False
    keyA = dataToList.key
    for rowData in TargetData.dataList:
        if keyA == rowData.key:
            hasSame = True
            Ddata = -1
            for data in dataToList.dataList:
                if data.dataStyle == 'D':
                    Ddata = float(data.SData)

            if Ddata == -1:
                print("ERR: combineSameNew Ddata == -1")

            for data in rowData.dataList:
                if data.dataStyle == 'D':
                    DdataT = float(data.SData) + Ddata
                    data.SData = str(DdataT)

        if hasSame == True:
            break
    return hasSame

def getSrcDataForRow(srcSheet, row, styleList, dataToList):
    for datainList in styleList.dataList:  # ����ѭ���б��ó�����������
        data = copy.deepcopy(datainList)  # ���ﲻ�ܰ����÷ŵ��б�����һ��Ҫ���

        if data.sClo != -1:
            data.SData = srcSheet.cell(row, data.sClo).value

            # �õ�һ�е�key
            dataToList.key = getKey(dataToList, data)

            dataToList.dataList.append(data)
    return dataToList


def getMapIdForRow(styleList, dataToList, idDict):
    for datainList in styleList.dataList:
        data = copy.deepcopy(datainList)  # ���ﲻ�ܰ����÷ŵ��б�����һ��Ҫ���

        if data.dataStyle == 'M':
            name = getName(dataToList.dataList)

            if name not in idDict:
                print("ERR: getMapIdForRow %s not in idDict" % name)

            id = idDict[name]  # ������Ҫ�쳣����
            data.SData = id
            dataToList.dataList.append(data)
            break
    return dataToList

def getSrcData(styleList, idDict, configdata):
    srcxlsname = configdata.get_srcxls("srcxlsname")
    srcRowStart = int(configdata.get_srcxls("srcRowStart"))
    endrowpara = configdata.get_srcxls("endrowpara")

    srcSheet = xlrd.open_workbook(srcxlsname).sheets()[0]
    stylerows = srcSheet.nrows

    TargetData = alldata()

    for row in range(srcRowStart, stylerows):
        para = srcSheet.cell(row, 0).value
        #print('para   = ', para)
        if para != endrowpara:
            dataToList = Rowdata().make_struct(row)

            dataToList = getSrcDataForRow(srcSheet, row, styleList, dataToList)

            # ӳ����������ȥ
            dataToList = getMapIdForRow(styleList, dataToList, idDict)

            hasSame = combineSame(TargetData, dataToList)

            if hasSame == False:
                TargetData.dataList.append(dataToList)
        else:
            break
    return TargetData

def writedata(TargetData, rowStart = 4):
    workbook = xlwt.Workbook(encoding='gbk')
    worksheet = workbook.add_sheet('Sheet1')

    for dataList in TargetData.dataList:
        for data in dataList.dataList:
            if data.dataStyle != 'D':
                worksheet.write(rowStart, data.dClo, data.SData)
            else:
                worksheet.write(rowStart, data.dClo, float(data.SData))
        rowStart = rowStart + 1

    workbook.save('�����ڴ�.xls')
    print("��� �����ڴ�.xls")

if __name__ == '__main__':
    configdata = ReadConfig()

    idDict =getMapIdByDict(configdata)

    styleList = Rowdata().make_struct()
    styleList = getStyle(configdata)

    TargetData = alldata()  # �������ݱ�
    TargetData = getSrcData(styleList, idDict, configdata)

    writedata(TargetData)

    print("�밴�س����˳�")
    input()


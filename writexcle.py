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
        def __init__(self, rownum = -1):
            self.dataList = []
            self.rownum = rownum

    def make_struct(self, rownum = -1):
        return self.Struct(rownum)

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

def getMapId(configdata):
    mapxlsname = configdata.get_map("mapxlsname")
    RowStart = int(configdata.get_map("RowStart"))
    nameclos = int(configdata.get_map("nameclos"))
    mapidclos = int(configdata.get_map("mapidclos"))

    srcbook = xlrd.open_workbook(mapxlsname)
    srcSheet = srcbook.sheets()[0]
    sumrows = srcSheet.nrows

    mapList = []
    for row in range(RowStart, sumrows):
        map = mapdata()
        map.name = srcSheet.cell(row, nameclos).value
        map.id = srcSheet.cell(row, mapidclos).value
        mapList.append(map)
    return mapList

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
        # print('cell.value = ', cell.value)
        if cell.value != 'MAP' :
            data.sClo = ord(cell.value) - ord('A')
        # print('data.sClo = ', data.sClo)
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

def getId(name, mapList):
    for map in mapList:
        if map.name == name:
            return map.id
    return ''


def getSrcData(styleList, mapList, configdata):
    srcxlsname = configdata.get_srcxls("srcxlsname")
    srcRowStart = int(configdata.get_srcxls("srcRowStart"))
    endrowpara = configdata.get_srcxls("endrowpara")

    srcbook = xlrd.open_workbook(srcxlsname)
    srcSheet = srcbook.sheets()[0]
    stylerows = srcSheet.nrows

    TargetData = alldata()

    for row in range(srcRowStart, stylerows):
        para = srcSheet.cell(row, 0).value
        #print('para   = ', para)
        if para != endrowpara:
            dataToList = Rowdata().make_struct(row)

            for datainList in styleList.dataList:
                data = copy.deepcopy(datainList)  # �б�ŵ��б�һ��Ҫ���

                if data.sClo != -1:
                    data.SData = srcSheet.cell(row, data.sClo).value
                    # print('data.sData  = ', data.SData )
                    dataToList.dataList.append(data)

            # ӳ����������ȥ
            for datainList in styleList.dataList:
                data = copy.deepcopy(datainList)  #һ��Ҫ���

                if data.dataStyle == 'M':
                    name = getName(dataToList.dataList)
                    id = getId(name, mapList)
                    data.SData = id
                    # print('data.sData  = ', data.SData )
                    dataToList.dataList.append(data)
                    break
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

if __name__ == '__main__':
    configdata = ReadConfig()

    mapList = []
    mapList = getMapId(configdata)

    styleList = Rowdata().make_struct()
    styleList = getStyle(configdata)

    TargetData = alldata()  # �������ݱ�
    TargetData = getSrcData(styleList, mapList, configdata)

    writedata(TargetData)


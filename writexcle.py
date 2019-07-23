#!/usr/bin/python3
#coding=gbk
import copy
import xlrd
import xlwt
from common.readconfig import ReadConfig

# 数据的最小单元格式
class DataInfo:
    """终表数据结构"""
    class Struct(object):
        def __init__(self, sClo = -1, dClo = -1, dataStyle = 'S', SData = ''):
            self.sClo = sClo            # 数据源列
            self.dClo = dClo            # 数据目的列
            self.dataStyle = dataStyle  # 数据格式
            self.SData = SData          # 字符型数据

        def make_struct(self, sClo, dClo, dataStyle, SData):
            return self.Struct(sClo, dClo, dataStyle, SData)
# 行数据格式
class Rowdata(object):
    """原始数据表"""
    class Struct(object):
        def __init__(self, rownum = -1):
            self.dataList = []
            self.rownum = rownum

    def make_struct(self, rownum = -1):
        return self.Struct(rownum)

class mapdata:
    """映射表"""
    name = ''
    id = ''


# 汇总数据格式
class alldata:
    """汇总数据表"""
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
                data = copy.deepcopy(datainList)  # 列表放到列表一定要深拷贝

                if data.sClo != -1:
                    data.SData = srcSheet.cell(row, data.sClo).value
                    # print('data.sData  = ', data.SData )
                    dataToList.dataList.append(data)

            # 映射表内容填进去
            for datainList in styleList.dataList:
                data = copy.deepcopy(datainList)  #一定要深拷贝

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

    workbook.save('数据在此.xls')

if __name__ == '__main__':
    configdata = ReadConfig()

    mapList = []
    mapList = getMapId(configdata)

    styleList = Rowdata().make_struct()
    styleList = getStyle(configdata)

    TargetData = alldata()  # 最终数据表
    TargetData = getSrcData(styleList, mapList, configdata)

    writedata(TargetData)


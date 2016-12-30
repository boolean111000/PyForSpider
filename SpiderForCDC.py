# coding:utf-8
# author:CuiCheng
# FinishTime: 2016/12/14
import urllib.request
from lxml import html
import xlwt
import copy
import datetime

def ExcelCount():
    CNT=[0]
    def add_one():
        CNT[0]=CNT[0]+1
        return CNT[0]
    return add_one
CNT = ExcelCount()

def GenerateUrl():

    years = ['2014', '2015', '2016']
    months = [str(i) for i in range(1, 13)]
    hrefs = []
    domin = 'http://www.gzcdc.org.cn/ajax/Report.aspx?data='
    for y in years:
        for m in months:
            href = domin + y + '-' + m.zfill(2)
            hrefs.append(href)
    return hrefs


def getLinks():
    TotalUrls = []
    hrefs = GenerateUrl()
    for href in hrefs:
        webheader = {'User-Agent': 'Mozilla/5.0 (Windows NT 6.1; WOW64; rv:23.0) Gecko/20100101 Firefox/23.0'}
        req = urllib.request.Request(href, headers=webheader)
        req = urllib.request.urlopen(req)
        response = req.read().decode('utf-8').strip()
        req.close()
        if (response != ''):
            print(href[-7:]+'有数据。')
            tree = html.fromstring(response)
            hrefs = tree.xpath('//a/@href')
            for href in hrefs:
                TotalUrls.append(href)
    print('共有'+str(len(TotalUrls))+'个数据表')
    return TotalUrls


def getInfo(url):
    webheader = {'User-Agent': 'Mozilla/5.0 (Windows NT 6.1; WOW64; rv:23.0) Gecko/20100101 Firefox/23.0'}
    req = urllib.request.Request(url, headers=webheader)
    req = urllib.request.urlopen(req)
    response = req.read().decode('utf-8')
    req.close()
    # Get the XML of the response
    root = html.fromstring(response)
    ExcelName = root.xpath('//h1/text()')[0]
    tables = root.xpath('//table')

    SheetsData = []
    for table in tables:
        SheetName = table.xpath('./preceding-sibling::p[position()<5]'
            '/strong/text()')
        SheetName = ''.join(SheetName).strip()
        if(SheetName == ''):
            SheetName = table.xpath('./preceding-sibling::p[2]//text()')
            SheetName = ''.join(SheetName).strip()
        # get the table header.
        Headers = []
        th = [[attr for attr in td.xpath('p//text()')] for td in table.xpath('tbody/tr[1]/td')]
        for eachitem in th:
            eachitem = ''.join(eachitem)
            Headers.append(eachitem)

        # get the rowspan and content
        rowspan = []
        district = []
        TotalData = []
        rows = table.xpath('tbody/tr')

        CountOfFields = len(rows[1])
        for eachitem in rows[1:]:
            Data = []
            if len(eachitem) == CountOfFields:
                data = eachitem.xpath('td[position()>1]/p//text()')
                pdata = eachitem.xpath('td[position()>1]')
                for eachpiece in pdata:
                    xiaodata = eachpiece.xpath('.//text()')
                    xiaodata = ''.join(xiaodata)
                    Data.append(xiaodata)

                row = eachitem.xpath('td[1]')[0]
                district.append(row.xpath('p//text()')[0])
                if row.get('rowspan') is not None:
                    rowspan.append(row.get('rowspan'))
                else:
                    rowspan.append(1)
            else:
                pdata = eachitem.xpath('td')
                for eachpiece in pdata:
                    xiaodata = eachpiece.xpath('.//text()')
                    xiaodata = ''.join(xiaodata)
                    Data.append(xiaodata)
                # data = eachitem.xpath('td/p//text()')

            TotalData.append(Data)
        SheetData = []
        SheetData.append(SheetName)
        SheetData.append(Headers)
        SheetData.append(rowspan)
        SheetData.append(district)
        SheetData.append(TotalData)
        SheetsData.append(SheetData)

    writeToExcel(ExcelName, SheetsData)


def writeToExcel(ExcelName, SheetsData):
    COUNT_OF_TABLE = CNT()
    print(COUNT_OF_TABLE)
    wbk = xlwt.Workbook(encoding='utf-8')
    for SheetData in SheetsData:
        SheetName = SheetData[0]
        Headers = SheetData[1]
        rowspan = SheetData[2]
        district = SheetData[3]
        TotalData = SheetData[4]
        writeToSheet(wbk, SheetName, Headers, rowspan, district, TotalData)
    print(ExcelName)
    wbk.save(str(COUNT_OF_TABLE).zfill(3) + ExcelName + '.xls')
    print(str(COUNT_OF_TABLE).zfill(3) + ExcelName + '处理完毕。')


def writeToSheet(wbk, SheetName, Headers, rowspan, district, TotalData):
    print(SheetName)
    sheet1 = wbk.add_sheet(SheetName, cell_overwrite_ok=True)
    # set the width of Table
    for i in range(len(Headers)):
        if(i != 0):
            sheet1.col(i).width = 256 * 22

    # ----------------
    # 设置表格样式（列宽和居中）

    style1 = xlwt.XFStyle()
    style1.alignment.horz = style1.alignment.HORZ_CENTER
    style1.alignment.vert = style1.alignment.VERT_CENTER
    bds = xlwt.Borders()
    bds.left = xlwt.Borders.THIN
    bds.right = xlwt.Borders.THIN
    bds.top = xlwt.Borders.THIN
    bds.bottom = xlwt.Borders.THIN
    style1.borders = bds

    style2 = copy.deepcopy(style1)
    style2.font.bold = True
    style2.font.name = 'Times New Roman'
    style2.font.colour_index = 2
    # ----------------
    # write the header of the table
    for i in range(0, len(Headers)):
        sheet1.write(0, i, Headers[i], style2)
    # write the first col of the table
    rowpointer = 1
    for i in range(0, len(rowspan)):
        sheet1.write_merge(rowpointer, rowpointer + int(rowspan[i]) - 1, 0, 0, district[i], style1)
        rowpointer = rowpointer + int(rowspan[i])
    # wirte the data
    rowoffset = 1
    coloffset = 1
    for i in range(len(TotalData)):
        for j in range(len(TotalData[i])):
            sheet1.write(i + rowoffset, j + coloffset, TotalData[i][j], style1)


def timeDecorator(func):
    def Wrapper():
        start = datetime.datetime.now()
        func()
        end = datetime.datetime.now()
        print('开始时间是： ' + str(start))
        print('结束时间是： ' + str(end))
        print('运行时长是： ' + str(end - start))
    return Wrapper


@timeDecorator
def main():
    basedomin = 'http://www.gzcdc.org.cn'
    urls = getLinks()
    for url in urls:
        getInfo(basedomin + url)

if __name__ == "__main__":
    main()

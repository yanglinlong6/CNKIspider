import xlsxwriter as xw
import pandas as pd

def xw_toExcel(data, fileName):  # xlsxwriter库储存数据到excel
    workbook = xw.Workbook(fileName)  # 创建工作簿
    worksheet1 = workbook.add_worksheet("sheet1")  # 创建子表
    worksheet1.activate()  # 激活表
    title = ['序号', '酒店', '价格']  # 设置表头
    worksheet1.write_row('A1', title)  # 从A1单元格开始写入表头
    i = 2  # 从第二行开始写入数据
    for j in range(len(data)):
        insertData = [data[j]["id"], data[j]["name"], data[j]["price"]]
        row = 'A' + str(i)
        worksheet1.write_row(row, insertData)
        i += 1
    workbook.close()  # 关闭表


def pd_toExcel(data, fileName):  # pandas库储存数据到excel
    ids = []
    names = []
    prices = []
    for i in range(len(data)):
        ids.append(data[i]["id"])
        names.append(data[i]["name"])
        prices.append(data[i]["price"])

    dfData = {  # 用字典设置DataFrame所需数据
        '序号': ids,
        '酒店': names,
        '价格': prices
    }
    df = pd.DataFrame(dfData)  # 创建DataFrame
    df.to_excel(fileName, index=False)  # 存表，去除原始索引列（0,1,2...）


if __name__ == '__main__':
    testData = [
        {"id": 1, "name": "立智", "price": 100},
        {"id": 2, "name": "维纳", "price": 200},
        {"id": 3, "name": "如家", "price": 300},
    ]
    fileName = '测试.xlsx'
    # xw_toExcel(testData,fileName)
    pd_toExcel(testData,fileName)
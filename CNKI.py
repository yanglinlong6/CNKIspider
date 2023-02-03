import os
import re
import time

from selenium.webdriver import Chrome
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
from tqdm import tqdm
from selenium.webdriver.common.by import By
import xlsxwriter as xw
import pandas as pd

# 存放搜索的关键字
global search


# 判断是否为中文摘要
def Chinese(web):
    try:
        web.find_element(By.XPATH, '//*[@id="ChDivSummary"]')
        return True
    except:
        return False


# 判断是否存在下一页
def Next_page(web):
    try:
        web.find_element(By.XPATH, '//*[@id="PageNext"]')
        cur = web.find_element(By.XPATH, '//*[@class="pages"]/span[@class="cur]').text
        # 这里控制搜索的页数
        if cur >= 2:
            return False
        else:
            return True
    except:
        return False


# 判断是否为英文摘要
def English(web):
    try:
        web.find_element(By.XPATH, '//*[@id="doc-summary-content-text"]')
        return True
    except:
        return False


# 判断是否打开了子网页
def is_childpage(web):
    try:
        web.switch_to.window(web.window_handles[1])
        return True
    except:
        return False


# 爬取数据
def spider(web):
    # key_search = input('请输入索引关键字：')
    # 获取所有成分，存放到medical_list中
    medical_list = get_data()
    print(medical_list)
    # 已抓取过的文章存放到列表have_search_list = []
    have_search_list = []
    # 用一个变量来标明第一次搜索和后面几次搜索
    count_mark = 0
    for key_search in medical_list:
        # 给全局变量赋值
        global search
        search = key_search
        # 如果已经抓取过，则跳过，进行下一个
        if key_search in have_search_list:
            continue
        else:
            have_search_list.append(key_search)
            # 每个成分创建一个文件夹
            file_path = './Abstract/' + key_search
            if not os.path.exists(file_path):
                os.mkdir(file_path)

        if count_mark == 0:
            web.find_element(By.XPATH, '//*[@id="txt_SearchText"]').send_keys(key_search, Keys.ENTER)
            count_mark += 1
        else:
            web.find_element(By.XPATH, '//*[@id="txt_search"]').clear()
            web.find_element(By.XPATH, '//*[@id="txt_search"]').send_keys(key_search, Keys.ENTER)
        # input('筛选是否完成：')
        # 筛选年份
        time.sleep(4)
        Choose_year(web)
        time.sleep(4)
        # while循环实现翻页
        page_count = 0

        datatables = []
        while True:
            try:
                time.sleep(2)
                # 获取论文列表
                page_count += 1
                tr_list = web.find_elements(By.XPATH, """//*[@id="gridTable"]/table/tbody/tr""")
                print(f'{key_search} : 正在读取第{page_count}页!')
                for tr in tqdm(tr_list):
                    time.sleep(1)
                    # 主题
                    title = tr.find_element(By.XPATH, """.//td[@class="name"]""").text
                    # 作者
                    author = tr.find_element(By.XPATH, """.//td[@class="author"]""").text
                    # 来源
                    source = tr.find_element(By.XPATH, """.//td[@class="source"]""").text
                    # 日期
                    date = tr.find_element(By.XPATH, """.//td[@class="date"]""").text
                    # 期刊
                    data = tr.find_element(By.XPATH, """.//td[@class="data"]""").text
                    # 下载次数
                    downloadCnt = tr.find_element(By.XPATH, """.//td[@class="download"]""").text

                    # 点击论文链接
                    tr.find_element(By.XPATH, './td/a[@class="fz14"]').click()
                    # 切换到新窗口
                    if is_childpage(web) == True:
                        web.switch_to.window(web.window_handles[1])
                    else:
                        print('*' * 100)
                        print('未打开新窗口')
                        time.sleep(50)
                        continue
                    time.sleep(3)
                    # 中文期刊
                    if Chinese(web) == True:
                        # 抓取标题
                        title = web.find_element(By.XPATH,
                                                 '/html/body/div[@class="wrapper"]/div[@class="main"]/div[@class="container"]/div[@class="doc"]/div[@class="doc-top"]/div[@class="brief"]/div[@class="wx-tit"]/h1').text
                        print('title', title)
                        author = web.find_element(By.XPATH,
                                                  '/html/body/div[@class="wrapper"]/div[@class="main"]/div[@class="container"]/div[@class="doc"]/div[@class="doc-top"]/div[@class="brief"]/div[@class="wx-tit"]/h3').text
                        print('author', author)
                        # 抓取摘要
                        abstract = web.find_element(By.XPATH, '//*[@id="ChDivSummary"]').text
                    # 英文期刊
                    elif English(web) == True:
                        # 抓取标题
                        title = web.find_element(By.XPATH, '//*[@id="doc-title"]').text
                        print('title', title)
                        author = ''
                        # 抓取摘要
                        abstract = web.find_element(By.XPATH, '//*[@id="doc-summary-content-text"]').text
                    else:
                        title = "标题不存在！"
                        author = ''
                        abstract = ''

                    # 如果标题中存在/则用汉字替换
                    r = r"[.!+-=——,$%^，,。？?、~@#￥%……&*《》<>「」{}【】()/\\\[\]'\"]"
                    title_copy = re.sub(r, ' ', title)
                    # title_copy = re.sub(r'/','斜杠',title)

                    # print(title)!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
                    # print(abstract)
                    # 写入excel表格
                    datatables.append({"title": title, "author": author, "source": source, "date": date, "data": data,
                                       "downloadCnt": downloadCnt, "abstract": abstract})
                    print('datatable', {"title": title, "author": author, "source": source, "date": date, "data": data,
                                        "downloadCnt": downloadCnt, "abstract": abstract})
                    # 写入摘要
                    try:
                        with open(file_path + '/result.txt', 'a', encoding='utf-8') as f:
                            f.write('\n')
                            f.write(title)
                            f.write('\n')
                            f.write(author)
                            f.write('\n')
                            f.write(abstract)
                            f.write('\n')
                            f.write('------------------------------------------------------------------------------')
                        if is_childpage(web) == True:
                            # 关闭子网页
                            web.close()
                            # 切换到原网页
                            web.switch_to.window(web.window_handles[0])
                        else:
                            web.switch_to.window(web.window_handles[0])
                    except:
                        print('文件名称Invalid !')
                        if is_childpage(web) == True:
                            # 关闭子网页
                            web.close()
                            # 切换到原网页
                            web.switch_to.window(web.window_handles[0])
                        else:
                            web.switch_to.window(web.window_handles[0])
                if Next_page(web) == True:
                    # web.find_element('//*[@id="PageNext"]').click()
                    web.find_element(By.XPATH, '/html/body').send_keys(Keys.RIGHT)
                else:
                    break
                time.sleep(3)
            except:
                print('定位失败！当前在第{}页！'.format(page_count))
                # 跳回第一个页面
                if is_childpage(web) == True:
                    web.switch_to.window(web.window_handles[1])
                    web.close()
                    web.switch_to.window(web.window_handles[0])
                else:
                    web.switch_to.window(web.window_handles[0])
                # 点击下一页
                if Next_page(web) == True:
                    # web.find_element('//*[@id="PageNext"]').click()
                    web.find_element(By.XPATH, '/html/body').send_keys(Keys.RIGHT)
                    time.sleep(5)
                else:
                    break
        # 写入Excel表格
        xw_toExcel(datatables, key_search)
        # pd_toExcel(datatables, key_search)
        print('药物：', key_search, 'Successful!')


# 筛选中文、年份
def Choose_year(web):
    count = 0
    while True:
        try:
            time.sleep(10)
            # 点击中文
            web.find_element(By.XPATH, '/html/body/div[3]/div[1]/div/div/div/a[1]').click()
            time.sleep(8)
            web.find_element(By.XPATH, '//*[@id="divGroup"]/dl[3]/dt/i[1]').click()
            time.sleep(8)
            web.find_element(By.XPATH, '//*[@id="divGroup"]/dl[3]/dt/i[2]').click()
            time.sleep(8)
            web.find_element(By.XPATH, '//*[@id="txtStartYear"]').send_keys(2015)
            time.sleep(3)
            web.find_element(By.XPATH, '//*[@id="txtEndYear"]').send_keys(2023)
            time.sleep(3)
            web.find_element(By.XPATH, '//*[@id="btnFilterYear"]').click()
            break
        except:
            web.refresh()
            # 重新输入
            web.find_element(By.XPATH, '//*[@id="txt_search"]').clear()
            time.sleep(2)
            web.find_element(By.XPATH, '//*[@id="txt_search"]').send_keys(search, Keys.ENTER)
            count += 1
            if count > 15:
                time.sleep(120)
            if count > 21:
                break


# 从本地csv文件中读取成分
def get_data():
    medical = []
    medical.append(input("请输入"))
    print('文章读取完成！')
    return medical


def xw_toExcel(data, fileName):  # xlsxwriter库储存数据到excel
    print('开始写表格',data)
    workbook = xw.Workbook(fileName + '.xlsx')  # 创建工作簿
    worksheet1 = workbook.add_worksheet("sheet1")  # 创建子表
    worksheet1.activate()  # 激活表
    title = ['主题', '作者', '来源', '日期', '期刊', '下载次数', '摘要']  # 设置表头
    worksheet1.write_row('A1', title)  # 从A1单元格开始写入表头
    i = 2  # 从第二行开始写入数据
    for j in range(len(data)):
        insertData = [data[j]["title"], data[j]["author"], data[j]["source"], data[j]["date"], data[j]["data"],
                      data[j]["downloadCnt"], data[j]["abstract"]]
        row = 'A' + str(i)
        worksheet1.write_row(row, insertData)
        i += 1
    workbook.close()  # 关闭表


def pd_toExcel(data, fileName):  # pandas库储存数据到excel
    print('开始写表格',data)
    titles = []
    authors = []
    sources = []
    dates = []
    datas = []
    downloadCnts = []
    abstracts = []
    for i in range(len(data)):
        titles.append(data[i]["title"])
        authors.append(data[i]["author"])
        sources.append(data[i]["source"])
        dates.append(data[i]["date"])
        datas.append(data[i]["data"])
        downloadCnts.append(data[i]["downloadCnt"])
        abstracts.append(data[i]["abstract"])

    dfData = {  # 用字典设置DataFrame所需数据
        '主题': titles,
        '作者': authors,
        '来源': sources,
        '日期': dates,
        '期刊': datas,
        '下载次数': downloadCnts,
        '摘要': abstracts
    }
    df = pd.DataFrame(dfData)  # 创建DataFrame
    df.to_excel(fileName + '-1.xlsx', index=False)  # 存表，去除原始索引列（0,1,2...）


if __name__ == '__main__':
    if not os.path.exists('./Abstract'):
        os.mkdir('./Abstract')

    opt = Options()
    opt.add_experimental_option('excludeSwitches', ['enable-automation'])
    opt.add_argument('--headless')
    opt.add_argument(
        'user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/92.0.4515.131 Safari/537.36')
    url = 'https://www.cnki.net/?pageId=77825&wfwfid=145305&websiteId=58201'
    web = Chrome(options=opt)
    web.get(url)
    # 开始爬取数据
    spider(web)
    # while True:
    #     spider(web)
    #     choice = input('是否搜索下一个关键字（y/n）：')
    #     if choice == 'n':
    #         break

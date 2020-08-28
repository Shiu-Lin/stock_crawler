# -*- coding: utf-8 -*-

from selenium import webdriver
import time
from bs4 import BeautifulSoup
from selenium.webdriver.support.ui import Select
import xlrd
import xlwt
import os
from xlutils.copy import copy
def write_excel(sheet_name,data):
    file_path = "./stock_info.xls"
    print sheet_name
    if os.path.isfile(file_path):#如果檔案存在，則打開，新增
        excel = xlrd.open_workbook(file_path,formatting_info=True)
        wb = copy(excel)
    else:#如果檔案不存在，直接新增一個新的excel
        wb = xlwt.Workbook()
    sheet = wb.add_sheet(sheet_name)
    x=0#行數
    for i in data:
        i = list(i)
        for r in range(0, len(i)):#列
            if len(i) == 8:
                sheet.write(x, r + 1, i[r].get_text())
            else:
                sheet.write(x, r, i[r].get_text())
        x+=1
    wb.save(file_path)


def data_crawler(index,sheet_name):
    wd = webdriver.Chrome("./chromedriver")
    wd.get("https://www.sitca.org.tw/ROC/Industry/IN2630.aspx?pid=IN22601_05/")
    time.sleep(5)
    select = Select(wd.find_element_by_id(("ctl00_ContentPlaceHolder1_ddlQ_Comid")))
    select.select_by_index(index)  # 下拉式選單選項選擇
    wd.find_element_by_xpath('//*[@id="ctl00_ContentPlaceHolder1_BtnQuery"]').click()
    # wd.refresh() #重新整理
    time.sleep(10)
    page_source = wd.page_source
    soup = BeautifulSoup(page_source, "html.parser")
    content = soup.find("td", {"class": "DTeven"})
    tbody = content.find_parent('tbody')  # 找到td的上層tbody，為了一次找到全部tr，因為tr的class 都不一樣。直接找全部tr會包含前他不需要的部分。
    tr = tbody.findAll('tr')  # 找tbody底下全部的tr
    write_excel(sheet_name, tr)
    wd.quit()

def get_bank_name():
    wd = webdriver.Chrome("./chromedriver")
    wd.get("https://www.sitca.org.tw/ROC/Industry/IN2630.aspx?pid=IN22601_05/")
    time.sleep(5)
    page_source = wd.page_source
    soup = BeautifulSoup(page_source, "html.parser")
    bank_list = soup.find("select", {"id": "ctl00_ContentPlaceHolder1_ddlQ_Comid"})
    all_bank = bank_list.findAll('option')  # 拿到全部下拉式選單選項
    wd.quit()
    return all_bank
if __name__=='__main__':
    all_bank = get_bank_name()
    index = 0
    for i in all_bank:
        bank_name = i.text.split()[1]
        data_crawler(index, bank_name)
        index += 1
    print "success!"


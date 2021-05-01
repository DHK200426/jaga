from selenium import webdriver
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.common.alert import Alert
from openpyxl import load_workbook
import time
#made by 1301 all rights reserved
#use selenium, openpyxl
def jaga():
    #open userdata and read
    load_wb = load_workbook("./inform.xlsx", data_only=True)
    sheet1 = load_wb['Sheet1']
    namelist = []
    birlist = []
    passlist = []
    for i in sheet1.rows:
        name = i[0].value
        namelist.append(name)
    for i in sheet1.rows:
        bir = i[1].value
        birlist.append(bir)
    for i in sheet1.rows:
        pass1 = i[2].value
        passlist.append(pass1)
    i = 0
    #save as namelist, birlist, passlist
    browser = webdriver.Chrome()
    
    while i <= len(namelist) - 1:
        browser.get("https://hcs.eduro.go.kr/#/loginHome")
        time.sleep(3)
        #school select
        browser.delete_all_cookies()
        browser.find_element_by_id("btnConfirm2").click()
        browser.find_element_by_class_name("searchBtn").click()
        Select(browser.find_element_by_id("sidolabel")).select_by_value("03")
        Select(browser.find_element_by_id("crseScCode")).select_by_value("4")
        scb = browser.find_element_by_id("orgname")
        time.sleep(0.2)
        scb.send_keys("대구일과학고등학교")
        scb.send_keys(Keys.RETURN)
        time.sleep(1)
        #user input
        browser.find_element_by_xpath('//*[@id="softBoardListLayer"]/div[2]/div[1]/ul/li/a').click()
        browser.find_element_by_xpath('//*[@id="softBoardListLayer"]/div[2]/div[2]/input').click()
        browser.find_element_by_xpath('//*[@id="user_name_input"]').send_keys(namelist[i])
        browser.find_element_by_xpath('//*[@id="birthday_input"]').send_keys(birlist[i])
        browser.find_element_by_xpath('//*[@id="btnConfirm"]').click() 
        time.sleep(2)

        browser.find_element_by_xpath('//*[@id="WriteInfoForm"]/table/tbody/tr/td/input').send_keys(passlist[i])
        browser.find_element_by_xpath('//*[@id="btnConfirm"]').click()
        time.sleep(3)
        browser.find_element_by_xpath('//*[@id="container"]/div/section[2]/div[2]/ul/li[1]/a/span[1]').click()
        time.sleep(2)
        try:
            #servay input
            browser.find_element_by_xpath('//*[@id="survey_q1a1"]').click()
            browser.find_element_by_xpath('//*[@id="survey_q2a1"]').click()
            browser.find_element_by_xpath('//*[@id="survey_q3a1"]').click()
            time.sleep(0.5)
            browser.find_element_by_xpath('//*[@id="btnConfirm"]').click()
            browser.find_element_by_xpath('//*[@id="topMenuBtn"]').click()
            time.sleep(0.5)
            browser.find_element_by_xpath('//*[@id="topMenuWrap"]/ul/li[4]/button').click()
            Alert(browser).accept()
            time.sleep(0.5)
            browser.find_element_by_xpath('/html/body/app-root/div/div[1]/div/button').click()
            Alert(browser).accept()
            print(namelist[i] + " 자가진단 완료")
            i = i + 1
        except:
            time.sleep(0.3)
            browser.find_element_by_xpath('//*[@id="topMenuBtn"]').click()
            time.sleep(0.5)
            browser.find_element_by_xpath('//*[@id="topMenuWrap"]/ul/li[4]/button').click()
            Alert(browser).accept()
            time.sleep(0.5)
            browser.find_element_by_xpath('/html/body/app-root/div/div[1]/div/button').click()
            Alert(browser).accept()
            print(namelist[i] + " 자가진단 본인이 완료")
            i = i + 1
    print("자가진단 완료")
    browser.close()
jaga()
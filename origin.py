#-*- coding:utf-8 -*-

from selenium import webdriver
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.common.alert import Alert
import time, datetime
import schedule
from pytz import timezone
import pytz
from random import * 
from openpyxl import load_workbook
from selenium.webdriver.chrome.options import Options
options = Options()
options.add_argument('--headless')
options.add_argument('--disable-gpu')
options.add_argument('--no-sandbox')
options.add_argument("--disable-dev-shm-usage")
#made by 1301 all rights reserve
#use selenium, openpyxl

#random start time 7:20~7:40
print("자가진단 프로세스 시작")

def jaga():
    #open userdata and read
    load_wb = load_workbook("./inform.xlsx", data_only=True)
    sheet1 = load_wb['Sheet1']
    namelist = []
    birlist = []
    passlist = []
    list = []
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
    browser = webdriver.Chrome(executable_path='/usr/lib/chromium-browser/chromedriver', options=options)

    #random user pick
    for k in range(len(namelist)):
            a = randint(0,len(namelist)-1) 
            while a in list:
                a =randint(0,len(namelist)-1)
            list.append(a)
    print(list)

    while i <= len(namelist) - 1:
        #random timedealy each person
        t = list[i]
        timedealy = randint(10,30)

        browser.get("https://hcs.eduro.go.kr/#/loginHome")
        time.sleep(3)
        #school select
        browser.delete_all_cookies()
        browser.find_element_by_id("btnConfirm2").click()
        browser.find_element_by_class_name("searchBtn").click()
        Select(browser.find_element_by_id("sidolabel")).select_by_value("03")
        Select(browser.find_element_by_id("crseScCode")).select_by_value("4")
        scb = browser.find_element_by_id("orgname")
        time.sleep(4)
        scb.send_keys("대구일과학고등학교")
        scb.send_keys(Keys.RETURN)
        time.sleep(4)
        #user input
        browser.find_element_by_xpath('//*[@id="softBoardListLayer"]/div[2]/div[1]/ul/li/a').click()
        browser.find_element_by_xpath('//*[@id="softBoardListLayer"]/div[2]/div[2]/input').click()
        browser.find_element_by_xpath('//*[@id="user_name_input"]').send_keys(namelist[t])
        browser.find_element_by_xpath('//*[@id="birthday_input"]').send_keys(birlist[t])
        time.sleep(5)
        browser.find_element_by_xpath('//*[@id="btnConfirm"]').click() 
        time.sleep(5)
        browser.find_element_by_xpath('//*[@id="WriteInfoForm"]/table/tbody/tr/td/input').send_keys(passlist[t])
        browser.find_element_by_xpath('//*[@id="btnConfirm"]').click()
        time.sleep(8)

        CheckPoint = browser.find_element_by_id('//*[@id="container"]/div/section[2]/div[2]/ul/li')

        if 'active' not in CheckPoint.get_attribute('class'):
            time.sleep(4)
            browser.find_element_by_xpath('//*[@id="topMenuBtn"]').click()
            time.sleep(4)
            browser.find_element_by_xpath('//*[@id="topMenuWrap"]/ul/li[4]/button').click()
            Alert(browser).accept()
            time.sleep(4)
            browser.find_element_by_xpath('/html/body/app-root/div/div[1]/div/button').click()
            Alert(browser).accept()
            i = i + 1
            print(namelist[t] + " 자가진단 본인이 완료" + str(i))
        
        browser.find_element_by_xpath('//*[@id="container"]/div/section[2]/div[2]/ul/li[1]/a/span[1]').click()
        time.sleep(5)
        browser.find_element_by_xpath('//*[@id="survey_q1a1"]').click()
        browser.find_element_by_xpath('//*[@id="survey_q2a1"]').click()
        browser.find_element_by_xpath('//*[@id="survey_q3a1"]').click()
        time.sleep(4)
        browser.find_element_by_xpath('//*[@id="btnConfirm"]').click()
        browser.find_element_by_xpath('//*[@id="topMenuBtn"]').click()
        time.sleep(4)
        browser.find_element_by_xpath('//*[@id="topMenuWrap"]/ul/li[4]/button').click()
        Alert(browser).accept()
        time.sleep(4)
        browser.find_element_by_xpath('/html/body/app-root/div/div[1]/div/button').click()
        Alert(browser).accept()
        i = i + 1
        print(namelist[t] + " 자가진단 완료" + str(timedealy) + " " + str(i))
        time.sleep(timedealy)
    print("자가진단 완료"+ str(nowDay))
    browser.close()
schedule.every().day.at('08:20').do(jaga)
while True:
    nowDay = time.strftime('%D')
    schedule.run_pending()
    time.sleep(1)

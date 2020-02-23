from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as ec
import os
import openpyxl
import time

dirpath = os.getcwd()
chromePath = dirpath+'\chromedriver\chromedriver.exe'

vk = openpyxl.Workbook()
sh = vk.active

huayClub_user = 'sub@zjjjc'
huayClub_pass = 'Aa112233'

SBO_user = 'subsbosbo'
SBO_pass = 'Aa112233+'


def huayClub():
    huayClub_driver = webdriver.Chrome(chromePath)
    huayClub_driver.maximize_window()
    huayClub_driver.set_page_load_timeout(10)
    print('HuayClub')
    try:     
        huayClub_driver.get('https://agent.superlot999.com/login')
        elem = huayClub_driver.find_element_by_name('username')
        elem.send_keys(huayClub_user)
        elem = huayClub_driver.find_element_by_name('password')
        elem.send_keys(huayClub_pass)
        elem = WebDriverWait(huayClub_driver, 10).until(ec.visibility_of_element_located((By.XPATH, '//*[@id="login-box"]/div/div/form/fieldset/div[2]/button')))
        elem.click()
        elem = WebDriverWait(huayClub_driver, 10).until(ec.visibility_of_element_located((By.XPATH, '//*[@id="sidebar-menu"]/li[1]/a')))
        elem.click()
        elem = WebDriverWait(huayClub_driver, 10).until(ec.visibility_of_element_located((By.XPATH, '//*[@id="sidebar-menu"]/li[1]/ul/li[2]/a')))
        elem.click()
        elem = WebDriverWait(huayClub_driver, 10).until(ec.visibility_of_element_located((By.ID, 'week')))
        elem.click()
        elem = huayClub_driver.find_element_by_xpath('//*[@id="game-type-list"]/div[2]/button')
        elem.click()
        sumAgent =  WebDriverWait(huayClub_driver, 10).until(ec.visibility_of_element_located((By.XPATH, '//*[@id="by-member-table-2"]/tfoot/tr/td[10]'))).text
        sumCom = WebDriverWait(huayClub_driver, 10).until(ec.visibility_of_element_located((By.XPATH, '//*[@id="by-member-table-2"]/tfoot/tr/td[14]'))).text
        print('Sum Agent = ' ,sumAgent)
        print('Sum Company = ' ,sumCom)
        sh['A19'].value='HUAYCLUB'
        sh['B19'].value=sumAgent
        sh['I19'].value='HUAYCLUB'
        sh['J19'].value=sumCom
        print('==============================')
    except:
        print("Something Error")


def saveToExcel():
    try:
        print("Saving Excel...")
        #Header Agent
        sh['A1'].value='ZJJJCG8'
        sh['A2'].value='Profit'
        sh['A3'].value='Web-User'
        sh['A20'].value='ค่าใช้จ่าย'
        sh['A21'].value='Promotion'
        sh['A22'].value='Promotion สมากชิกใหม่'
        sh['A23'].value='ยอดยกมา'
        sh['A26'].value='Total-ProFit'

        sh['B3'].value='จำนวน'
        sh['D2'].value='คำนวน'
        sh['E2'].value='คำนวน'

        #Header Company
        sh['G1'].value='Company'
        sh['G2'].value='Profit'
        sh['G3'].value='Web-User'
        sh['G20'].value='ค่าใช้จ่าย'
        sh['G23'].value='ยอดยกมา'
        sh['G26'].value='Total-Company'
    
        sh['H3'].value='จำนวน'
        sh['J2'].value='คำนวน'
        sh['K2'].value='คำนวน'

        #Header Summary
        sh['A28'].value='สรุปยอด'
        sh['B28'].value='จำนวน'
        sh['A29'].value='Office 60%'
        sh['A30'].value='Partner 40%'
        sh['A31'].value='company'

        vk.save(dirpath+'\output\Report.xlsx')
        print("Save Excel Success!!")
    except:
        print("Save excel Error")
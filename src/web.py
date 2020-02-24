#!/usr/bin/env python
# -*- coding: utf-8 -*- #
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as ec
from openpyxl import load_workbook
import os 
import ConfigParser
import datetime


dirpath = os.getcwd()
chromePath = dirpath+'\chromedriver\chromedriver.exe'
phantom = dirpath+'\\phantomjs\\bin\\phantomjs.exe'
phantomjsPath = str.replace(phantom,'\\','\\\\')

config = ConfigParser.RawConfigParser()
config.read(dirpath+'\ConfigFile.properties')

def huayClub():
    print '=============================='
    print '>>> HuayClub <<<<'
    
    try:
        
        huayClub_url = config.get('HuayClubSection', 'huayClub_url')
        huayClub_user = config.get('HuayClubSection', 'huayClub_user')
        huayClub_pass = config.get('HuayClubSection', 'huayClub_pass')
        print '---------------'
        print 'PhantomJs is runing...'
        driver = webdriver.PhantomJS(executable_path = phantomjsPath)
        print '---------------'
        driver.maximize_window()
        print 'Now Hitting... '+huayClub_url
        driver.get(huayClub_url)
    
        elem = driver.find_element_by_name('username')
        elem.send_keys(huayClub_user)
        elem = driver.find_element_by_name('password')
        elem.send_keys(huayClub_pass)
        elem = WebDriverWait(driver, 10).until(ec.visibility_of_element_located((By.XPATH, '//*[@id="login-box"]/div/div/form/fieldset/div[2]/button')))
        elem.click()
        elem = WebDriverWait(driver, 10).until(ec.visibility_of_element_located((By.XPATH, '//*[@id="sidebar-menu"]/li[1]/a')))
        elem.click()
        elem = WebDriverWait(driver, 10).until(ec.visibility_of_element_located((By.XPATH, '//*[@id="sidebar-menu"]/li[1]/ul/li[2]/a')))
        elem.click()
        elem = WebDriverWait(driver, 10).until(ec.visibility_of_element_located((By.ID, 'week')))
        elem.click()
        elem = driver.find_element_by_xpath('//*[@id="game-type-list"]/div[2]/button')
        elem.click()

        agent = WebDriverWait(driver, 10).until(ec.visibility_of_element_located((By.XPATH, '//*[@id="by-member-table-2"]/tbody/tr/td[1]'))).text    
        sumAgent =  WebDriverWait(driver, 10).until(ec.visibility_of_element_located((By.XPATH, '//*[@id="by-member-table-2"]/tbody/tr/td[11]'))).text
        sumCom = WebDriverWait(driver, 10).until(ec.visibility_of_element_located((By.XPATH, '//*[@id="by-member-table-2"]/tbody/tr/td[15]'))).text
        
        driver.close()
        print'Sum Agent = ' ,agent
        print'Sum Agent = ' ,sumAgent
        print 'Sum Company = ' ,sumCom

        date_object = datetime.date.today()
        wb = load_workbook(dirpath+'\output\Report.xlsx')
        print '[HuayClub] Opening Excel File...'
        ws = wb['Sheet1']
        ws['B2'].value=date_object
        ws['C19'].value=agent
        ws['B19'].value=sumAgent
        ws['H19'].value=sumCom
        
        wb.save(dirpath+'\output\Report.xlsx')
        print '[HuayClub] Save Excel Success!!'
        print '=============================='
    except:
        print "Something Error"


def writeExcel(sheet, role1, role2, value1, value2):
    print 'Hello'
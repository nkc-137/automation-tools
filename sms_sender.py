# -*- coding: utf-8 -*-
"""
Created on Sat Jun 20 11:28:15 2015

@author: Kapil Chandra
"""
import xlrd
from selenium.webdriver import Firefox
import time
from selenium import webdriver 
#opening the xl sheet which is having all the numbers
workbook= xlrd.open_workbook('sms_sender.xlsx')
#going into the particular sheet of the xl file
worksheet = workbook.sheet_by_name('Sheet1')
#finding number of num_rows in the sheet
num_rows = worksheet.nrows -1
#opening firefox browser
driver = webdriver.Firefox()
#going into way2sms portal
driver.get("http://www.way2sms.com")
#the website has an advertisement displaying for 4-5 seconding so I sleep my code
time.sleep(20)
#filing my mobile number and password to login
inputElement = driver.find_element_by_name("username")
inputElement.send_keys("8500409855")
inputElement = driver.find_element_by_name("password")
inputElement.send_keys("28091997")
inputElement.submit()
#the site is slower so I sleep it for a while
time.sleep(5)
#the below line press the button Send Free Sms
driver.find_element_by_xpath("//input[@class='button br3'][@value='Send Free SMS']").click()
#this path locates send sms and press the button
driver.find_element_by_xpath("/html/body/div[7]/div[1]/ul/li[2]/a").click()
#Now the site goes into an html(inner) so I switched the frame
driver.switch_to.frame(driver.find_element_by_id('frame'))
#filling number and message to whom we have send
inputElement = driver.find_element_by_xpath("/html/body/div[3]/div[1]/form/div[2]/div[1]/input")
inputElement.send_keys(str(8500409855))
inputElement = driver.find_element_by_xpath("/html/body/div[3]/div[1]/form/div[2]/textarea")
inputElement.send_keys("sucessful")
driver.find_element_by_name("Send").click()
#the above was a test message to user and the loop takes the numbers from xl and sends respective messages
for i in range(0,num_rows+1):
    time.sleep(2)
    driver.find_element_by_xpath("/html/body/form/div[1]/div[1]/p[1]").click()
    inputElement = driver.find_element_by_xpath("/html/body/div[3]/div[1]/form/div[2]/div[1]/input")
    inputElement.send_keys(str(worksheet.cell_value(i,0)))
    go =driver.find_element_by_xpath("/html/body/div[3]/div[1]/form/div[2]/textarea")
    go .send_keys("susessful")
    driver.find_element_by_name("Send").click()
    
    
    

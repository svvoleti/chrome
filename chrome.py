import os
import time
from unicodedata import name
import openpyxl
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium import webdriver
obj = webdriver.Chrome(r"C:\Users\91991\Desktop\chromedriver.exe")
obj.get("https://www.flipkart.com/")
obj.implicitly_wait(10)
wb=openpyxl.load_workbook()
sheet1=wb["sheet1"]
for i in range(2, sheet1.max_row+1):
    keyword=sheet1['A'+str(i)].value
    obj.find_element(BY.XPATH,"/html/body/div[1]/div/div[1]/div/div/div/div/div[1]/div/div/div/div[1]/div[1]/header/div[1]/div[2]/form/div/div/input").send_keys(keyword)
    obj.find_element(BY.XPATH,"/html/body/div[1]/div/div[1]/div/div/div/div/div[1]/div/div/div/div[1]/div[1]/header/div[1]/div[2]/form/div/div/input").send_keys(keys.ENTER)
    cat=obj.find_element(BY.XPATH,"/html/body/div/div/div[3]/div[1]/div[1]/div/div[1]/div/div/section/div[2]/div/a").text
    obj.find_element("/html/body/div[1]/div/div[1]/div/div/div/div/div[1]/div/div/div/div[1]/div[1]/header/div[1]/div[2]/form/div/div/input").send_keys(keys.CONTROL,'a')
    obj.find_element("/html/body/div[1]/div/div[1]/div/div/div/div/div[1]/div/div/div/div[1]/div[1]/header/div[1]/div[2]/form/div/div/input").send_keys(keys.BACKSPACE)
    last=obj.execute_script("/html/body/div/div/div[3]/div[1]/div[1]/div/div[1]/div/div/section/div[2]/div/a").innertext
    sheet1['c'+str(i)].value=last
    wb.save("Book1.xlsx")








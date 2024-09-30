from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
import time
import pandas as pd


chr_options = Options()
chr_options.add_experimental_option("detach", True)
driver = webdriver.Chrome(options=chr_options)

def Welcome():
    print("#######################################################")
    print("# Selenium StocksBot programs (SSBp) Version 4.0      #")
    print("# This project (SSBp) was developed by M_inuyashiki   #")
    print("#######################################################")
    print("###########################")
    print("# Welcome to SSBp_Project #")
    print("#   Plaese choose menu    #")
    print("#   1.Get Price Stocks    #")
    print("#   2.Get ImportentsData  #")
    print("#   3.Exit Program        #")
    print("###########################")

def FindData():
    try:
        Number = int(input("Number : "))
        for i in range(Number):
            Stockname = str(input("Symbolstock " + str(i+1) + " : "))
            Stocknamelist.append(Stockname)

        Name = str(input("Filename : "))
        for i in range(Number):
            if Stocknamelist != None:
                driver.get("https://www.set.or.th/th/market/product/stock/quote/"+ Stocknamelist[i] +"/price")
                Symbol = driver.find_element(By.XPATH, '//*[@id="__layout"]/div/div[2]/div[1]/div[5]/div[2]/div[2]/div[2]/h1').text
                Price = driver.find_element(By.XPATH, '//*[@id="__layout"]/div/div[2]/div[1]/div[5]/div[2]/div[2]/div[3]/div[2]').text
                Change = driver.find_element(By.XPATH, '//*[@id="__layout"]/div/div[2]/div[1]/div[5]/div[2]/div[2]/div[3]/h3/span[1]').text
                Change_per = driver.find_element(By.XPATH, '//*[@id="__layout"]/div/div[2]/div[1]/div[5]/div[2]/div[2]/div[3]/h3/span[2]').text
                Max = driver.find_element(By.XPATH,'//*[@id="stock-quote-tab-pane-1"]/div/div[1]/div/div[2]/div[1]/div/div[2]/div[2]/div[2]/div[1]/span').text
                Min = driver.find_element(By.XPATH,'//*[@id="stock-quote-tab-pane-1"]/div/div[1]/div/div[2]/div[1]/div/div[2]/div[2]/div[2]/div[2]/span').text
                Date = driver.find_element(By.XPATH,'//*[@id="stock-quote-tab-pane-1"]/div/div[1]/div/div[2]/div[1]/div/div[1]/div[2]/span').text
            else:
                print("Plaese Add Symbol")
            for i in range(1):
                a = {"ชื่อหุ้น" : Symbol,
                    "ราคา" : Price,
                    "การเปลี่ยนแปลง" : Change,
                    "เปลี่ยนแปลง %" : Change_per,
                    "ราคาสูงสุด" : Max,
                    "ราคาตํ่าสุด" : Min,
                    "ข้อมูล ณ วันที่ " : Date,
                    }
                Stocks.append(a)
    except ValueError:
        print("Can you get label not number ?")
    Name = Name+".xlsx"   
    # print(Stocks)
    df = pd.DataFrame(Stocks).to_excel(Name,index=False)
    print("บันทึกข้อมูลสำเร็จ")
    driver.quit()

time.sleep(2)

Stocks = []
Stocknamelist = []

Welcome()
Option = str(input(""))
if Option == "1":
    FindData()
elif Option == "2":
    print("Not support in V4")
    exit()
else:
    print("Good Bye bro !!!")
    exit()




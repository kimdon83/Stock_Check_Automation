# %% import os
import logging
import time
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.alert import Alert

from bs4 import BeautifulSoup
import requests
import pandas as pd
import json

import os

start = time.time()

# get the current working directory
current_dir = os.getcwd()

# Configure logging settings
# log_filename = current_dir+"my_log_file.log"
# logging.basicConfig(filename=log_filename, filemode="a", level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s")

# %%
desktoploc=r'C:\Users\dokim2\OneDrive - Kiss Products Inc\Desktop'
with open(desktoploc+'\IVYENT_DH\data.json', 'r') as f:
    data = json.load(f)

driver = webdriver.Chrome()
driver.maximize_window()  # 창 최대화

driver.get(data['okta_url'])

# %%
time.sleep(5)
login=driver.find_element(By.ID,"input28")
login.click()
login.send_keys(data['okta_id'])

time.sleep(1)
login=driver.find_element(By.ID,"input36")
login.click()
login.send_keys(data['okta_pw'])
driver.find_element(By.CLASS_NAME,"button").click() # click login button

time.sleep(5)
driver.get(data['portal_url'])

# %%
import pyautogui
time.sleep(1)

pyautogui.write(data['okta_id'])
time.sleep(0.1)
pyautogui.press('tab')
time.sleep(0.1)
pyautogui.write(data['okta_pw'])
time.sleep(0.1)
pyautogui.press('tab')
time.sleep(0.1)
pyautogui.press('enter')

time.sleep(3)

# a=input("you have to login to okta before going to sample request")

requestID=input("samplerequest id:")

aaa=driver.get(f'http://portal.kissusa.com/Lists/Sample%20Request/DispForm.aspx?ID={requestID}')

# abb=driver.find_element(By.ID,"titl186-2_")
# acc=abb.find_element(By.CLASS_NAME,"ms-commentexpand-iconouter")
# acc.click()
# table=driver.find_element(By.ID,"tbod186-2__")
df=pd.read_html(driver.page_source)[0]
df0=df[19:-25]
df0.columns=['SalesOrg','material','ip','qty','plant','Sloc','stock','ms','actCov','openOrd','OpenDelivery','AsnDate','ASNqty','ASNcov']
df1=df0[['material','plant','qty']]


desktoploc=r'C:\Users\dokim2\OneDrive - Kiss Products Inc\Desktop'
resultLoc=desktoploc+"\stock check practice"
# input_df.to_csv(resultLoc+"\\"+"simulator_input.csv",index=False) #Change Location - (type 2)

df1.insert(3,"order_number",requestID)
df1.insert(4,"salesorg",1100)
df1.to_csv(resultLoc+"\\"+"simulator_input.csv",index=False)

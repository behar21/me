import webbrowser
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.common.exceptions import TimeoutException
from bs4 import BeautifulSoup
import time
from selenium.webdriver.common.keys import Keys
import openpyxl
from selenium.webdriver.support.ui import Select

#-------------- Starts: Load Excel ------------------------
wb = openpyxl.Workbook()
wb = openpyxl.load_workbook(filename='policies copy.xlsm')
sheets = wb.sheetnames
ws = wb[sheets[0]]
print("excel has been loaded")
#-------------- Ends: Load Excel ------------------------

#-------------- Starts: Web Driver ----------------------
#driver = webdriver.Chrome()
#driver.set_window_size(1124, 850) 
#driver.get("https://app.smile-direct.ch/car?lang=de&sac=WOA")
#wait = WebDriverWait(driver, 3)
#-------------- Ends: Web Driver ------------------------

#try:
# ------------- Starts: Reading Excel -------------------
brand = ws['P2'].value
model = ws['T2'].value+" "+ws['U2'].value
inverhrkersetzung = ws['v2'].value
accessories = ws['R2'].value
leasing = ws['H2'].value
gender = ws['J2'].value
use = ws['I2'].value
licenceAge = ws['K2'].value
nat = ws['L2'].value
bd = '01.01.1976'
deductibleTeil = ws['N2'].value
deductibleVoll = ws['O2'].value
if licenceAge != '5+':
    bd == '01.01.1996'

zipcode = ws['M2'].value

licence = '1996'
if licenceAge != '5+':
    licence = '2015'
print("Attributes initialised")
#-------------- Ends: Reading Excel ---------------------
#--------------- Stars: first submission page ------------------------------------
#wait.until(EC.element_to_be_clickable((By.FOR,'markeundtyp_radio')))

print("loop ready to start")
for index in range(1,100):
    driver = webdriver.PhantomJS()
    driver.set_window_size(1024, 768)
    wait = WebDriverWait(driver, 10)
    driver.get("https://app.smile-direct.ch/car?lang=de&sac=WOA")
    
    print("Driver has been initialized")
    try:
        driver.find_element_by_xpath("//label[@for='markeundtyp_radio']").click()
        
        wait.until(EC.presence_of_element_located((By.ID, 'brand')))
        driver.find_element_by_xpath("//*[@id='brand']/option[text()='"+brand+"']")
        
        wait.until(EC.visibility_of_element_located((By.ID, 'type')))
        driver.find_element_by_xpath("//*[@id='type']/option[text()='"+model+"']")
        
        print(index)
    except:
        print "Took too much time to load"
    finally:
        print("The End")
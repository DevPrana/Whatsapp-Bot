import xlrd
import time
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from pynput.keyboard import Key,Controller
#from selenium.webdriver.support.ui import WebDriverWait
#from selenium.webdriver.support import expected_conditions as EC
#from selenium.webdriver.common.by import By
from os import system
import pyautogui

keyboard=Controller()
adres="Whatsapp.xlsx"
wb=xlrd.open_workbook(adres)
sheet=wb.sheet_by_index(0)
numr=sheet.nrows                        #number of rows in excel sheet
print("WELCOME TO MY SOFTWARE")
txt=""
while len(txt)==0:
    txt=input("Please enter valid message to send to contacts: ")
opt=input("do you want to attatch something Y/N: ")
if opt=="Y" or opt=="y":
    opt2=int(input('''===========================
What do you want to attatch
press 1 for image
press 2 for contact
:- '''))
if opt2==1:
    img=input("Enter image name with extension:- ")
elif opt2==2:
    cont=input("Enter exact contact number:- ")
numc=sheet.ncols
driverpath="C:\chromedriver\chromedriver"
driver=webdriver.Chrome(driverpath)
driver.get("https://web.whatsapp.com/")
time.sleep(10)
if opt2==1:
    for i in range(numr):
        number=sheet.cell_value(i,0)
        number=int(number)
        elem =driver.find_element_by_xpath("//*[@id='side']/div[1]/div/label/div/div[2]")
        elem.send_keys(number)
        elem.send_keys("\n")
        elem=driver.find_element_by_xpath("//*[@id='main']/footer/div[1]/div[2]/div/div[2]")
        elem.send_keys(txt)
        elem.send_keys("\n")
        elem=driver.find_element_by_xpath("//*[@id='main']/header/div[3]/div/div[2]/div").click()
        time.sleep(2)
        elem=driver.find_element_by_xpath("//*[@id='main']/header/div[3]/div/div[2]/span/div/div/ul/li[1]").click()
        time.sleep(3)
        keyboard.type(img)
        keyboard.press(Key.enter)
        keyboard.release(Key.enter)
        time.sleep(3)
        keyboard.press(Key.enter)
        keyboard.release(Key.enter)
elif opt2==2:
    for i in range(numr):
        number=sheet.cell_value(i,0)
        number=int(number)
        elem =driver.find_element_by_xpath("//*[@id='side']/div[1]/div/label/div/div[2]")  
        elem.send_keys(number)
        elem.send_keys("\n")
        elem=driver.find_element_by_xpath("//*[@id='main']/footer/div[1]/div[2]/div/div[2]")
        elem.send_keys(txt)
        elem.send_keys("\n")
        elem=driver.find_element_by_xpath("//*[@id='main']/header/div[3]/div/div[2]/div").click()
        time.sleep(2)
        elem=driver.find_element_by_xpath("//*[@id='main']/header/div[3]/div/div[2]/span/div/div/ul/li[4]").click()
        time.sleep(2)
        keyboard.type(cont)
        time.sleep(2)
        keyboard.press(Key.enter)
        keyboard.release(Key.enter)
        time.sleep(2)
        send_butt=pyautogui.locateCenterOnScreen('Send.png')
        pyautogui.click(send_butt)
        time.sleep(2)
        send_butt=pyautogui.locateCenterOnScreen('Send.png')
        pyautogui.click(send_butt)
        

    
            
        
    
    

    

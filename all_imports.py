import win32gui
import sys
import re
import win32con
import win32com.client 
from time import sleep
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.desired_capabilities import DesiredCapabilities
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import pyautogui as pg
import os
from selenium.webdriver.common.alert import Alert
from shutil import move, copymode,make_archive
import zipfile
import subprocess
from datetime import date

today = date.today()
csv_path = f'D:\Plag\plag({today}).csv'


def getVars(line):
    with open('D:\Plag\Vars.txt','r') as f:
        return f.readlines()[line].removesuffix('\n')
        
try:
    try:
        hWnd = win32gui.FindWindow(None, 'Turnitin - Google Chrome')
        win32gui.BringWindowToTop(hWnd)
    except Exception:
        hWnd = win32gui.FindWindow(None, 'Turnitin - Class Portfolio - Google Chrome')
        win32gui.BringWindowToTop(hWnd)
except Exception as e:
    print(e)
    subprocess.call(f'"D:\Chrome_(8989).bat"') 
    print('=========== Starting {0}.bat ==================='.format(getVars(5)))
 

capa = DesiredCapabilities.CHROME
capa["pageLoadStrategy"] = "none"
opt = Options()
opt.add_experimental_option("debuggerAddress",f"localhost:8989")
opt.add_argument('--headless')
opt.add_argument('--disable-gpu')
try:
    driver = webdriver.Chrome(executable_path=r'C:\Users\Priyam\.wdm\drivers\chromedriver\win32\103.0.5060.134\chromedriver.exe',options=opt)
except Exception:
    from webdriver_manager.chrome import ChromeDriverManager
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()),options=opt)

plgpath = r'D:\Plag\plagdir'
raw_path = r'D:\Plag\Raw'
block_cipher = None

def get_number_oflines(path):
    with open(r"{0}".format(path), 'r') as fp:
        count = 0
        for line in fp:
            if line != "\n":
                count += 1
    return count


def zipdir(path, ziph):
    # ziph is zipfile handle
    for root, dirs, files in os.walk(path):
        for file in files:
            if file.endswith('.pdf'):
                ziph.write(os.path.join(root, file),
                           os.path.relpath(os.path.join(root, file),
                                           os.path.join(path, '..')))
                try:
                    os.remove(f'{path}\\{file}')
                except Exception as e:
                    print(e)
                print(f'[Result]: {file} has been zipped')

def tab(n):
    tn = driver.window_handles[n]
    driver.switch_to.window(tn)

def Wait(xpth):
    return WebDriverWait(driver, 2).until(EC.presence_of_element_located((By.XPATH, xpth)))

def findxpth(xpath):
    try:
        return driver.find_element(by=By.XPATH,value=xpath)
    except Exception:
        pass

def getCsv(lines):#Return Any Line of plag.csv
    with open(csv_path,'r') as f:
        return f.readlines(0)[lines].split(',')[1].removesuffix('\n')

def plagfileslot(i):
    file_slot = driver.find_element(by=By.XPATH,value=f'/html/body/div[2]/div[4]/div[2]/div[2]/div[3]/table/tbody/tr[{i+1}]/td[1]').get_attribute('innerText')
    return file_slot

def extract_Zip():
    for dirpath, dirs, files in os.walk(raw_path):
            for filename in files:
                try:
                    if filename.endswith('.zip') or filename.endswith('.rar') :
                        file_name = f'{dirpath}\\{filename}'
                        zip_ref = zipfile.ZipFile(file_name) # create zipfile object
                        zip_ref.extractall(raw_path) # extract file to dir
                        zip_ref.close() # close file
                        os.remove(file_name) # delete zipped file
                        print('[INFO]',file_name,'Zip file extracted')
                    else:
                        pass
                except Exception as e:
                    print('Exception in extract zip: ',e)

def remove_emty_folder():
    folders = list(os.walk(raw_path))[1:]
    for folder in folders:
        # folder example: ('FOLDER/3', [], ['file'])
        if not folder[2]:
            os.rmdir(folder[0])
            
def move_all_doc():
    extract_Zip()
    for dirpath, dirs, files in os.walk(raw_path):
        for filename in files:
            if filename.endswith('.docx'):
                move(f'{dirpath}\\{filename}',f'{plgpath}\\{filename}')
        
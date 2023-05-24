from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
import time
from selenium.webdriver.support.select import Select
import pyautogui
import datetime
from selenium.webdriver.common.action_chains import ActionChains
import traceback
import sys
#import line_notify
import pyscreenshot as ImageGrab
from retrying import retry
from PIL import Image
import ddddocr
import cv2
import os
import logging
import pyperclip as pc
import re
from bs4 import BeautifulSoup

#@retry(stop_max_attempt_number = 2)
#初始值
account = ""#帳號
pwd = ""#密碼
url = 'https://www.tabf.org.tw/Service/Login.aspx'#網址
#日期設定
currentDateAndTime = datetime.datetime.now()#取得現在時間
currentDate = currentDateAndTime.strftime("%Y%m%d")#取得年月日，例：20230521
currentTime = currentDateAndTime.strftime("%H%M%S")#取得時間，例：154314
#資料夾&截圖設定
if not((os.path.exists(currentDate)) and (os.path.isdir(currentDate))): os.mkdir(currentDate)
os_dir = os.path.join(os.getcwd(), currentDate)
verification_pic = os_dir +'\\image.png'#驗證碼截圖
success_pic = os_dir +'\\'+ currentTime +'_success.png'#成功截圖
certnotinlist_pic = os_dir +'\\'+ currentTime +'_certnotinlist.png'#成功截圖
fail = os_dir +'\\'+currentTime+"_fail"+".png"#失敗進catch截圖
#log設定
log = os_dir + currentTime+"log.txt"
# create logger obj
logger = logging.getLogger()
# set log level
logger.setLevel(logging.INFO)
# file handler
handler = logging.FileHandler(log, 'w', encoding='utf-8')
handler.setFormatter(logging.Formatter("%(asctime)s-%(name)s-%(levelname)s: %(message)s"))
logger.addHandler(handler)
#後門
#TODO:測試時設定today日期
#後門－自行輸入日期更改判斷第幾個營業日，範例在後門/Test/Config.txt
if (os.path.isdir(os.path.join(os.getcwd(), 'Test'))) and (os.path.isfile(os.path.join(os.getcwd(),'Test', 'Config.txt'))):
    f = open((os.path.join(os.getcwd(),'Test', 'Config.txt')), 'r', encoding='utf-8')
    mylist = [line.rstrip('\n') for line in f]
    account = mylist[0]
    pwd = mylist[1]
    certifacation = mylist[2]
    area = mylist[3]
    print(account+'\n' + pwd+ '\n'+ area+ '\n'+ certifacation)
    logger.info('有後門資料，更改查詢資料為'+account+'\n' + pwd + '\n'+ area+ '\n'+ certifacation)
#後門
#驗證碼辨識
def screenshot_code_verificate(path):
    driver.save_screenshot(path)#為驗證碼先截全視窗圖
    element = driver.find_element("xpath",'//*[@id="ctl00_ctl00_ContentPlaceHolder1_ContentPlaceHolder1_Login1_imgCaptcha"]')
    left = element.location['x']
    top = element.location['y']
    #right = element.size['width']
    #bottom = element.size['height']
    right = element.location['x'] + element.size['width']
    bottom = element.location['y'] + element.size['height']
    image = Image.open(path)# 將 screen load 至 memory
    image = image.crop((left, top, right, bottom)) # 根據位置剪裁圖片
    image.save(path, 'png')               # 重新儲存圖片為 png 格式
    #驗證碼辨識
    ocr = ddddocr.DdddOcr()
    with open(path, 'rb') as f:
        img_bytes = f.read()
    res = ocr.classification(img_bytes)
    print(res)
    logger.info('驗證碼辨識結果為:'+res)
    return res
#從登入網站到選擇日期
def run():
    logger.info('啟動預定程式序號為'+'(0表示尚未取得當日訂位;1表示僅取得1900訂位;2表示僅取得2000訂位)')
    driver.implicitly_wait(240)
    driver.get(url)
    time.sleep(2)
    driver.maximize_window()
    time.sleep(2)
    #登入
    WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.XPATH, '//*[@id="wrapper"]/main/div[2]/div/div[1]/div[1]/div/span[1]/input[2]'))
    ).click()#帳號
    time.sleep(0.5)
    pc.copy(account)
    pyautogui.hotkey('ctrl', 'v')
    driver.find_element("xpath",'//*[@id="wrapper"]/main/div[2]/div/div[1]/div[2]/div/span[1]/input[2]').click() #密碼
    time.sleep(0.5)
    pc.copy(pwd)
    pyautogui.hotkey('ctrl', 'v')
    ans = screenshot_code_verificate(verification_pic)
    ans_int = int(ans[:1])+int(ans[2:3])
    ans = str(ans_int)
    print(ans)
    time.sleep(0.5)
    pc.copy(ans)
    driver.find_element("xpath",'//*[@id="wrapper"]/main/div[2]/div/div[1]/div[3]/div[2]/div[1]/input[2]').click() #驗證碼
    pyautogui.hotkey('ctrl', 'v')
    time.sleep(0.5)
    driver.find_element("xpath",'//*[@id="ctl00_ctl00_ContentPlaceHolder1_ContentPlaceHolder1_Login1_ibtnLogin"]').click() #登入鍵
    time.sleep(0.5)
    WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.XPATH, '//*[@id="wrapper"]/header/div[2]/div/ul[2]/li[3]/a'))
    ).click()#證照
    '''
    WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.XPATH, '//*[@id="wrapper"]/header/div[2]/div/ul[2]/li[3]/div/div[2]/ul/li[2]/a'))
    ).click()#證照搜尋
    '''
    time.sleep(2)
    if driver.find_element("xpath",'//*[@id="popupUIClose"]/button'):
        time.sleep(0.5)
        driver.find_element("xpath",'//*[@id="popupUIClose"]/button').click()
        logger.info('有公告訊息視窗，把她按掉')
    # 等待請選擇證照名稱下拉選單元件
    WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.XPATH, '//*[@id="ctl00_ContentPlaceHolder1_QuickSearch_LicenseSearchBar1_ddlTestName"]'))
    )
    #抓取請選擇證照名稱下拉選單元件
    text = driver.find_element("xpath",'//*[@id="ctl00_ContentPlaceHolder1_QuickSearch_LicenseSearchBar1_ddlTestName"]').text
    print(text)
    if certifacation not in text :
        driver.find_element("xpath",'//*[@id="ctl00_ContentPlaceHolder1_QuickSearch_LicenseSearchBar1_ddlTestName"]').click()
        driver.get_screenshot_as_file(certnotinlist_pic)
        raise ValueError(certifacation + 'is not in' + text)
    select = Select(driver.find_element("xpath",'//*[@id="ctl00_ContentPlaceHolder1_QuickSearch_LicenseSearchBar1_ddlTestName"]'))
    select.select_by_visible_text(certifacation)
    time.sleep(1)
    # 等待請選擇考區下拉選單元件
    WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.XPATH, '//*[@id="ctl00_ContentPlaceHolder1_QuickSearch_LicenseSearchBar1_ddlTestArea"]'))
    )
    #抓取請選擇考區下拉選單元件
    areatext = driver.find_element("xpath",'//*[@id="ctl00_ContentPlaceHolder1_QuickSearch_LicenseSearchBar1_ddlTestArea"]').text
    print(areatext)
    if area not in areatext :
        driver.find_element("xpath",'//*[@id="ctl00_ContentPlaceHolder1_QuickSearch_LicenseSearchBar1_ddlTestArea"]').click()
        driver.get_screenshot_as_file(certnotinlist_pic)
        raise ValueError(area + 'is not in' + areatext)
    select = Select(driver.find_element("xpath",'//*[@id="ctl00_ContentPlaceHolder1_QuickSearch_LicenseSearchBar1_ddlTestArea"]'))
    select.select_by_visible_text(area)
    time.sleep(1)
    driver.find_element("xpath",'//*[@id="ctl00_ContentPlaceHolder1_QuickSearch_btnSearch"]').click()
    time.sleep(1)
    
    #抓取當頁表格

    WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.XPATH, '//*[@id="wrapper"]/main/section/div/table'))
    )
    table = driver.find_element("xpath",'//*[@id="wrapper"]/main/section/div/table').text

    print(table)

    soup = BeautifulSoup(driver.page_source, "html5lib")
    id = soup.find_all('id')

    #抓取頁數表格並判斷最大頁數
    WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.XPATH, '//*[@id="ctl00_ContentPlaceHolder1_LicenseProductList1_RPT_Paginate"]'))
    )
    page = driver.find_element("xpath",'//*[@id="ctl00_ContentPlaceHolder1_LicenseProductList1_RPT_Paginate"]').text
    print(page)
    page = page.replace(' ','')
    #判斷最大頁數
    page=re.findall('\d+',page)
    print(page)
    if len(page)==1:
        page_max=int(''.join(page))%10
    else:
        # num_list_new = page[1]
        page_max=int(''.join(page[1]))
    print(page_max)

    #for i in range(int(page_max)):
        #print(str(i))


    return True
    '''
    #driver.find_element("xpath",'//*[@id="USERNAME"]').send_keys(account) #帳號
    '''
attemps = 0
success = False
while attemps < 3 and not success:#重複執行，最多錯3次
    try:
        driver = webdriver.Chrome()
        nexttime = run()#執行登入到選取預定日期
        logger.info('此次程式執行結果為'+'(0表示尚未取得當日訂位;1表示僅取得1900訂位;2表示僅取得2000訂位;12表示兩個時段都訂位成功)')
        success = nexttime
    except Exception as e:
        #img = ImageGrab.grab()
        attemps+=1
        print('錯誤發生，今日重新嘗試第'+str(attemps)+'次')
        print(e.__class__.__name__)#取得錯誤類型
        print(e.args[0])#取得詳細內容
        cl, exc, tb = sys.exc_info()
        lastCallStack = traceback.extract_tb(tb)[-1]#取得Call Stack的最後一筆資料
        print(lastCallStack[1])#取得發生的行號
        #img.save(filename, quality=70)
        driver.get_screenshot_as_file(fail)
        logger.error('錯誤發生，今日重新嘗試第'+str(attemps)+'次;')
        logger.error('錯誤行號為:'+str(lastCallStack[1])+';錯誤類行為:'+str(e.__class__.__name__)+';錯誤內容為:'+str(e.args[0]))
        driver.close()
        #if attemps == 3:
            #break
    finally:
        driver.quit()
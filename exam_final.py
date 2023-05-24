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

#option
chrome_options = webdriver.ChromeOptions()
chrome_options.add_argument('-enable-webgl')
chrome_options.add_argument('--no-sandbox')
chrome_options.add_argument('--disable-dev-shm-usage')

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
    print(area+ '\n'+ certifacation)
    logger.info('有後門資料，更改查詢資料為'+ area+ '\n'+ certifacation)
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
    logger.info('啟動程式')
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
    logger.info('驗證碼輸入值為'+ans)
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
    logger.info('登入成功並點選證照鍵')
    '''
    WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.XPATH, '//*[@id="wrapper"]/header/div[2]/div/ul[2]/li[3]/div/div[2]/ul/li[2]/a'))
    ).click()#證照搜尋
    '''
    time.sleep(1)
    WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.XPATH, '//*[@id="popupUIClose"]/button'))
    ).click()#證照
    logger.info('有公告訊息視窗，把她按掉')
    
    '''
    if driver.find_element("xpath",'//*[@id="popupUIClose"]/button'):
        time.sleep(0.5)
        driver.find_element("xpath",'//*[@id="popupUIClose"]/button').click()
    '''    
        
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
    logger.info('選擇查詢證照'+certifacation)
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
    logger.info('選擇考試區域'+area)
    driver.find_element("xpath",'//*[@id="ctl00_ContentPlaceHolder1_QuickSearch_btnSearch"]').click()
    time.sleep(1)
    logger.info('查詢開始')
    

    #soup = BeautifulSoup(driver.page_source, "html5lib")
    #htmlvalue = soup.find_all('td',string = '測驗時間')
    
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
    logger.info(f'取得頁數共{page_max}頁')

    if int(page_max) >= 1:
        for i in range(1,int(page_max)+1):
            #抓取當頁表格列數
            WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.XPATH, '//*[@id="wrapper"]/main/section/div/table'))
            )
            table = driver.find_element("xpath",'//*[@id="wrapper"]/main/section/div/table').text
            print(table)
            table_row = table.count(certifacation)
            print(table_row)
            logger.info(f'取得第{i}頁列數:'+ f'共{table_row}筆')
            for j in range(1,int(table_row)+1):
                test_time = driver.find_element("xpath",f'//*[@id="wrapper"]/main/section/div/table/tbody/tr[{j}]/td[4]').text
                print(test_time)
                logger.info(f'取得第{i}頁第{j}筆考試時間資料:'+ f'資料內容為{test_time}')
                if '六' in test_time or '日' in test_time:
                    apply_status = driver.find_element("xpath",f'//*[@id="wrapper"]/main/section/div/table/tbody/tr[{j}]/td[6]').text
                    print(apply_status)
                    logger.info(f'取得第{i}頁第{j}筆考試假日可取號狀態:'+ f'{apply_status}')
                    if apply_status == '可取號':
                        driver.find_element("xpath",f'//*[@id="ctl00_ContentPlaceHolder1_LicenseProductList1_RPT_Test_ctl0{j}_lbtnSend"]').click()
                        get_apply = 'ok'
                        logger.info(f'第{i}頁第{j}筆假日成功取號')
                        break
                    else:
                        get_apply = f'第{i}頁第{j}筆假日無法取號'
                        logger.info(get_apply)
                
                elif '18:50' in test_time:
                    apply_status = driver.find_element("xpath",f'//*[@id="wrapper"]/main/section/div/table/tbody/tr[{j}]/td[6]').text
                    print(apply_status)
                    logger.info(f'取得第{i}頁第{j}筆考試平日晚上可取號狀態:'+ f'{apply_status}')
                    if apply_status == '可取號':
                        driver.find_element("xpath",f'//*[@id="ctl00_ContentPlaceHolder1_LicenseProductList1_RPT_Test_ctl0{j}_lbtnSend"]').click()
                        get_apply = 'ok'
                        logger.info(f'第{i}頁第{j}筆平日晚上成功取號')
                        break
                    else:
                        get_apply = f'第{i}頁第{j}筆平日晚上無法取號'
                        logger.info(get_apply)
                else:
                    get_apply = '第'+str(i)+f'頁第{j}筆不符合假日與平日晚上'
                    logger.info(get_apply)
                    print(get_apply)
                    
            if get_apply == 'ok':
                break
            else:
                logger.info(f'第{i}頁無可取號場次')
            driver.find_element("xpath",'//*[@id="ctl00_ContentPlaceHolder1_LicenseProductList1_lbtnNextPageNum"]').click()
    else:
        #抓取當頁表格列數
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, '//*[@id="wrapper"]/main/section/div/table'))
        )
        table = driver.find_element("xpath",'//*[@id="wrapper"]/main/section/div/table').text
        print(table)
        table_row = table.count(certifacation)
        print(table_row)
        logger.info(f'此次查詢僅一頁共{table_row}筆')
        for j in range(1,int(table_row)+1):
            test_time = driver.find_element("xpath",f'//*[@id="wrapper"]/main/section/div/table/tbody/tr[{j}]/td[4]').text
            print(test_time)
            logger.info(f'取得第{j}筆考試時間資料:'+ f'資料內容為{test_time}')
            if '六' in test_time or '日' in test_time:
                apply_status = driver.find_element("xpath",f'//*[@id="wrapper"]/main/section/div/table/tbody/tr[{j}]/td[6]').text
                print(apply_status)
                logger.info(f'取得第{j}筆考試假日可取號狀態:'+ f'{apply_status}')
                if apply_status == '可取號':
                    driver.find_element("xpath",f'//*[@id="ctl00_ContentPlaceHolder1_LicenseProductList1_RPT_Test_ctl0{j}_lbtnSend"]').click()
                    get_apply = 'ok'
                    logger.info(f'第{j}筆假日成功取號')
                    break
                else:
                    get_apply = f'第{j}筆假日無法取號'
                    logger.info(get_apply)
            elif '18:50' in test_time:
                apply_status = driver.find_element("xpath",f'//*[@id="wrapper"]/main/section/div/table/tbody/tr[{j}]/td[6]').text
                print(apply_status)
                logger.info(f'取得第{j}筆考試平日晚上可取號狀態:'+ f'{apply_status}')
                if apply_status == '可取號':
                    driver.find_element("xpath",f'//*[@id="ctl00_ContentPlaceHolder1_LicenseProductList1_RPT_Test_ctl0{j}_lbtnSend"]').click()
                    get_apply = 'ok'
                    logger.info(f'第{j}筆平日晚上成功取號')
                    break
                else:
                    get_apply = f'第{j}筆平日晚上無法取號'
                    logger.info(get_apply)
            else:
                get_apply = f'第{j}筆不符合假日與平日晚上'
                logger.info(get_apply)
                print(get_apply)
    if get_apply == 'ok':
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, '//*[@id="ctl00_ContentPlaceHolder1_ExamDetail_btnGetTicket"]'))
        )

        bookok  = driver.find_element("xpath",'//*[@id="ctl00_ContentPlaceHolder1_ExamDetail_btnGetTicket"]')
        bookok.send_keys("\n")


        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, '//*[@id="ctl00_ContentPlaceHolder1_ExamDetail_imgbtnShopping"]'))
        )
        driver.get_screenshot_as_file(success_pic)
        logger.info(f'取號成功！報名時間為：{test_time}')
        return get_apply + test_time
    else:
        return 'fail'

attemps = 0
success = False
while attemps < 3 and not success:#重複執行，最多錯3次
    try:

        driver = webdriver.Chrome(options=chrome_options)
        nexttime = run()#執行登入到選取預定日期
        if  'ok' in nexttime:
            logger.info(f'此次程式執行結果為ok，取號成功，取號日期為{nexttime[2:]}')
        else:
            logger.info('程式執行完成，無可取號之場次！')
        success = True
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
#-*- coding: utf-8 -*-

import pyautogui
import webbrowser
import pyperclip
import timeunit
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from smtplib import SMTP_SSL
from selenium.common.exceptions import NoSuchElementException 
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from datetime import datetime
from time import sleep

from openpyxl import Workbook

search_keyword = '교육'

SMTP_SERVER = "smtp.naver.com"
SMTP_PORT = "465"
SMTP_USER = "pjhyl1127"
SMTP_PASSWORD = "zxczxc1127"

count = 0

def send_mail(name, conut, addr, contents, attachment=False) :
    msg = MIMEMultipart("alternative")

    if attachment :
        msg = MIMEMultipart('mixed')
    
    msg['From'] = SMTP_USER
    msg['To'] = addr
    msg['Subject'] = name+'님, '+ conut +'개의 나라장터 검색결과 입니다.'

    text = MIMEText(contents)
    msg.attach(text)

    if attachment:
        from email.mime.base import MIMEBase
        from email import encoders

        file_data = MIMEBase('application', 'octet-stream')
        f = open(attachment, 'rb')
        file_contents = f.read()
        file_data.set_payload(file_contents)
        encoders.encode_base64(file_data)

        from os.path import basename
        filename = basename(attachment)
        file_data.add_header('Content-Disposition', 'attachment', filename=filename)
        msg.attach(file_data)

    smtp = SMTP_SSL(SMTP_SERVER, SMTP_PORT)
    smtp.login(SMTP_USER, SMTP_PASSWORD)
    smtp.sendmail(SMTP_USER, addr, msg.as_string())
    smtp.close()

xlsx = Workbook()
sheet = xlsx.active
sheet.append(['Title', 'Open_date', 'End_date'])


now = datetime.now()
#today = str(now.year)+'/'+str(now.month)+'/'+str(now.day)
today = datetime.today().strftime("%Y/%m/%d") 

driver = webdriver.Chrome(r'C:\chromedriver.exe')

driver.get('http://www.g2b.go.kr/index.jsp')
sleep(0.5)

driver.switch_to.frame(driver.find_element_by_id('maintop_iframe'))
driver.find_element_by_class_name('keyword').send_keys(search_keyword)
serch_button = driver.find_element_by_id('AKCFrm')
serch_button.find_element_by_xpath('./fieldset/a').send_keys('\n')
#serachButton = pyautogui.locateCenterOnScreen('search.png')
#pyautogui.moveTo(serachButton)
#pyautogui.moveRel(-50,0)
#pyautogui.click(serachButton)
sleep(0.5)

# pyautogui.typewrite(search_keyword, interval=0.1)
# pyperclip.copy(search_keyword)
# pyautogui.hotkey('ctrl', 'v')
# pyautogui.press('enter')
# sleep(0.5)
pyautogui.press('enter')
sleep(3)

#element = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.NAME, "search_area")))
driver.switch_to_default_content()
driver.switch_to.frame(driver.find_element_by_name('sub'))
div = driver.find_element_by_class_name('search_area')
list_names = div.find_elements_by_xpath('./ul/li')
for list_name in list_names : 
    try :
        tilte = list_name.find_element_by_xpath('./strong/a')
        li_1 = list_name.find_element_by_class_name('m1')
        deadline = li_1.find_element_by_xpath('./span')
        li_2 = list_name.find_element_by_class_name('m2')
        open_date = li_2.find_element_by_xpath('./span')
        if(today in open_date.text) :
            sheet.append([tilte.text, open_date.text, deadline.text])
            count = count + 1
    except NoSuchElementException :
        tilte = list_name.find_element_by_xpath('./strong/a')
        li_2 = list_name.find_element_by_class_name('m2')
        open_date = li_2.find_element_by_xpath('./span')
        if(today in open_date.text) :
            sheet.append([tilte.text, open_date.text, "na"])
            count = count + 1


driver.quit()
file_name = 'naraMarket.xlsx'
xlsx.save(file_name)
    
send_mail("박정환", str(count),"pjhyl1127@naver.com",today+" 나라장터 검색 내역입니다.", file_name)
print(str(count))

     
# driver.switch_to_default_content()
# frame = driver.find_element_by_name('sub')
# div = frame.find_element_by_class_name('search_area')
# list_name = frame.find_elements_by_xpath('./ul/li/strong/a')
#print(list_name)
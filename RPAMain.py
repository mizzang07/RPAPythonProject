from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from datetime import datetime
import time

import pandas as pd

import smtplib
from email.mime.text import MIMEText
from email.header import Header

sDate = str(datetime.now().date()).replace('-', '')
print(sDate)
'''
options = webdriver.ChromeOptions()
options.add_argument("--start-maximized")
options.add_experimental_option("excludeSwitches", ["enable-logging"])
driver = webdriver.Chrome(options=options)
# 사이트 접속
driver.get("https://www.kebhana.com/cont/mall/mall15/mall1503/index.jsp")
time.sleep(2)
# 사이트 안 프레임으로 접속
driver.switch_to.frame('bankIframe')
# 기간환율변동 라디오박스 클릭
driver.find_element(By.XPATH , '//*[@id="inqFrm"]/table/tbody/tr[1]/td/span/p/label[3]').click()
start_box = driver.find_element(By.ID , 'tmpInqStrDt_p')
start_box.clear()
start_box.send_keys(sDate)
end_box = driver.find_element(By.ID , 'tmpInqEndDt_p')
end_box.clear()
end_box.send_keys(sDate)
# 통화선택
driver.find_element(By.ID , 'curCd').send_keys('USD')
# 고시회차 최초선택
driver.find_element(By.XPATH, '//*[@id="inqFrm"]/table/tbody/tr[6]/td/span/p/label[1]').click()
# 조회버튼 클릭
search_box = driver.find_element(By.XPATH, '//*[@id="HANA_CONTENTS_DIV"]/div[2]/a/span' ).click()
time.sleep(3)
# 엑셀다운버튼 클릭
driver.find_element(By.XPATH, '//*[@id="searchContentDiv"]/div[1]/a[2]/span').click()
time.sleep(10)


#엑셀 파일 읽어서 매매기준율 읽어오기
fileName = "C:/Users/LEE/Downloads/환율변동조회_" + sDate + ".xls"
print(fileName)

df = pd.read_excel(fileName, usecols=[7], header=2)
exchange_usa = str(df.iat[1,0])
print(exchange_usa)
'''
exchange_usa = "1568"
#네이버 메일 연결
smtp = smtplib.SMTP('smtp.naver.com', 587)
smtp.ehlo()
smtp.starttls()
smtp.login('whiteharu', 'gkfn7735!spdlqj')

#네이버 메일 보내기
my_mail = 'whiteharu@naver.com' #송신메일
your_mail = 'whiteharu@naver.com' #수신메일
subject = '오늘의 환율' + sDate #제목
message = "오늘의 환율은 " + exchange_usa + " 입니다."
msg = MIMEText(message.encode('utf-8'), _charset='utf-8')
msg['Subject'] = Header(subject.encode('utf-8'), 'utf-8')
msg['From'] = my_mail
msg['To'] = your_mail

smtp.sendmail(my_mail, your_mail, msg.as_string())
smtp.quit()
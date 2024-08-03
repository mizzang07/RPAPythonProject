from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from datetime import datetime

import time

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
radio_box = driver.find_element(By.XPATH , '//*[@id="inqFrm"]/table/tbody/tr[1]/td/span/p/label[3]').click()
start_box = driver.find_element(By.ID , 'tmpInqStrDt_p')
start_box.clear()
start_box.send_keys(str(datetime.now().date()))
end_box = driver.find_element(By.ID , 'tmpInqEndDt_p')
end_box.clear()
end_box.send_keys(str(datetime.now().date()))
# 통화선택
usa_box = driver.find_element(By.ID , 'curCd')
usa_box.send_keys('USD')
# 고시회차 최초선택
driver.find_element(By.XPATH, '//*[@id="inqFrm"]/table/tbody/tr[6]/td/span/p/label[1]').click()
# 조회버튼 클릭
search_box = driver.find_element(By.XPATH, '//*[@id="HANA_CONTENTS_DIV"]/div[2]/a/span' ).click()

# 엑셀다운버튼 클릭
driver.find_element(By.XPATH, '//*[@id="searchContentDiv"]/div[1]/a[2]/span').click()
time.sleep(10)


# 파파고(https://papago.naver.com)웹페이지를 이용하여,
# 문장을 입력하면 번역결과를 출력하는 프로그램.
# 1. 번역할 내용 입력:
# 2. 번역할 언어: (1)영어 (2)일본어 (3)중국어(간체)
# 해서 번역하는 프로그램을 코딩하세요.

# 버튼ㅣ button#ddTargetLanguageButton
#
from selenium import webdriver
import time

driver = webdriver.Chrome("./chromedriver")
driver.get("https://papago.naver.com/")
time.sleep(1) #지연시간 1초

content = input("번역할 내용: ")
language = input("번역할 언어: (1)영어 (2)일본어 (3)중국어(간체) ")
time.sleep(10) #지연시간 1초


language_list = driver.find_elements_by_css_selector("ul.dropdown_menu_inner___29_zc li.select_item___1U0X9")
# 언어 옵션이 리스트로 저장됨
# 0~45번 중에 영어[17] 일본어[18] 중국어간체[19]
if(language == 1): # 영어
    language_list[17].click()
if(language == 2): # 일본어
    language_list[18].click()
if(language == 3): # 중국어간체
    language_list[19].click()

time.sleep(1) #지연시간 1초

# 번역할 내용 입력: textarea#txtSource
txt = driver.find_element_by_css_selector("textarea#txtSource")
txt.send_keys(content)
# 버튼: button#btnTranslate
button = driver.find_element_by_css_selector("button#btnTranslate")
button.click()
time.sleep(5) #지연시간 1초

translate = driver.find_element_by_css_selector("div#txtTarget span").text
print(translate)

driver.close()

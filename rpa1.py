from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import requests
import time
import datetime
from slack import *
from config import get_secret
import pandas as pd
import pyautogui as gui

def nextXpath(path):
    return driver.find_element('xpath', path)

df = pd.read_excel("data/전화번호리스트.xlsx")


email = get_secret("email")
password = get_secret("pw")
# msg = "안녕하세요. 테스트입니다. 감사합니다."
url = "https://www.modusign.co.kr/"

driver = webdriver.Chrome(executable_path="./chromedriver/chromedriver.exe")

driver.get(url)
time.sleep(2)


# 로그인
nextXpath('//*[@id="Header"]/div/ul/li[7]/a[2]/b').click()
time.sleep(2)
nextXpath('//*[@id="root"]/div[2]/div/div/div/div/div[1]/form/div[1]/input').send_keys(email)
nextXpath('//*[@id="root"]/div[2]/div/div/div/div/div[1]/form/div[2]/div/input').send_keys(password)
nextXpath('//*[@id="root"]/div[2]/div/div/div/div/div[1]/form/button').click()
time.sleep(1)

results = []

for i in range(len(df['사번'])):
    # print(df.iloc[i].values)
    대상자 = df.iloc[i].values
    print(대상자)

    time.sleep(3)
    #팝업창닫기
    # nextXpath('//*[@id="getsitecontrol-314364"]//div/div/div/div[1]/img').click()  
    
    # 팝업창 닫기
    # print("총 윈도우 개수 {0}".format(len(driver.window_handles)))
    # pop_up = len(driver.window_handles)

    # for i in range(pop_up - 1):
    #     print("{0}번째 팝업창 닫기".format(i + 1))
    #     driver.switch_to.window(
    #         driver.window_handles[1])  # 팝업창 선택후 닫기 (팝업창의 개수는 변경될 수 있음)
    #     time.sleep(0.5)
    #     driver.close()
    #     time.sleep(0.5)

    # print("버튼찾기")
    # gui.moveTo(gui.locateCenterOnScreen('./button_x.png'))
    # gui.click()
    # print("버튼닫기")


    # time.sleep(1)
    #템플릿 화면 이동
    nextXpath('//*[@id="root"]/div[2]/div/div[2]/div[2]/section[1]/ul/li[3]/a/div/span').click()  
    time.sleep(1)
    
    # 첫번째 템플릿 선택
    nextXpath('//*[@id="root"]/div[2]/div/div[4]/div/div/div[2]/ul/div[1]/li/ul/li[1]/span').click()  
    time.sleep(2)

    #서명요청하기 클릭
    nextXpath('//*[@id="root"]/div[2]/div/div[4]/div/div/div[2]/ul/div[1]/li/div[2]/div/ul/li[1]').click()
    time.sleep(1)

    # 확인하고 시작하기 클릭
    nextXpath('/html/body/div[6]/div/div/div/div/div[2]/button[2]').click()   
    time.sleep(3)

    # 이름 또는 회사명 입력
    nextXpath('/html/body/div[1]/div[2]/div/div/div/div/div[2]/div/div/div[2]/div[1]/div[2]/input').click()
    nextXpath('/html/body/div[1]/div[2]/div/div/div/div/div[2]/div/div/div[2]/div[1]/div[2]/input').send_keys(str(대상자[1]))  
    time.sleep(1)

    # 카카오톡 전송버튼 클릭
    nextXpath('//*[@id="first-participant"]/div[1]/div[3]/button[2]').click() 
    time.sleep(1)

    # 전화번호 입력(뒤에 여덟자리만)
    nextXpath('/html/body/div[1]/div[2]/div/div/div/div/div[2]/div/div/div[2]/div[1]/div[3]/div/div[2]/div/div[1]/input').click()
    nextXpath('/html/body/div[1]/div[2]/div/div/div/div/div[2]/div/div/div[2]/div[1]/div[3]/div/div[2]/div/div[1]/input').send_keys(str(대상자[2])) 
    time.sleep(1)

    #다음 단계로 버튼 클릭
    nextXpath('//*[@id="root"]/div[2]/div/div/div/div/div[3]/div/div/button/span').click() 
    time.sleep(1)

    #다음 단계로 버튼 클릭
    nextXpath('//*[@id="root"]/div[2]/div/div/div/div/div[4]/div/div/button/span').click() 
    time.sleep(1)

    # 서명유효기간 설정 클릭
    nextXpath('/html/body/div[1]/div[2]/div/div/div/div/div[2]/div/div[1]/div/div[3]/p').click()  
    time.sleep(1)

    # 서명유효기간 박스 클릭
    nextXpath('//*[@id="root"]/div[2]/div/div/div/div/div[2]/div/div[1]/div/div[3]/div/input').clear()   
    time.sleep(1)

    # 서명유효기간 입력
    nextXpath('//*[@id="root"]/div[2]/div/div/div/div/div[2]/div/div[1]/div/div[3]/div/input').send_keys("7")   
    time.sleep(0.5)

    # nextXpath('//*[@id="root"]/div[2]/div/div/div/div/div[2]/div/div[4]/div/div[2]/button').click() # 남길말 입력 클릭
    # time.sleep(3)

    # nextXpath('/html/body/div[1]/div[2]/div/div/div/div/div[2]/div/div[4]/div/div[2]').send_keys(msg) # 남길말 입력하기

    # 설정완료 버튼 클릭
    nextXpath('//*[@id="root"]/div[2]/div/div/div/div/div[3]/div/button[2]/span').click()  
    time.sleep(2)

    # 개인정보수집 동의
    nextXpath('//*[@id="agreement-check1"]').click() 
    time.sleep(0.2)

    # 서명시 횟수차감 동의
    nextXpath('//*[@id="agreement-check2"]').click()  
    time.sleep(0.2)

    # 서명요청하기 클릭
    nextXpath('/html/body/div[6]/div/div/div/div/div[5]/div/button').click() 
    time.sleep(1)

    # 확인 버튼 클릭
    # nextXpath('/html/body/div[6]/div/div/div/div/button').click() 
    # time.sleep(3)
    msg = f"{i}번째 {대상자[1]} 발송 완료"
    print(msg)
    results.append(msg)
    
    if (int(i)+1) % 50 == 0:
        Slack_Msg(f"{int(i)+1}번째 발송 완료")
        
    # result_df = pd.DataFrame({"사번": 대상자[0] ,"결과": results})
    # result_df.to_excel("./results1.xlsx")
    # exit()

print("전체 발송 완료")
print(results)
result_df = pd.DataFrame({"사번": 대상자[0] ,"결과": results})
result_df.to_excel("./results1.xlsx")
Slack_Msg(f"전체 {int(i)+1}명 발송 완료")


# WebDriverWait(driver, 10).until(EC.presence_of_element_located(By.XPATH, "~~~")) # 기본 대기시간 설정하되, 로딩 완료시 바로 실행하는 코드

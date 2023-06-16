from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
from slack import *
from config import get_secret
import pandas as pd

def nextXpath(path):
    return driver.find_element('xpath', path)

def num_check(num):
    if num >= 8:
        return str(8)
    else:
        return str(7)    

def web_wait(경로):
    try:
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, 경로)))
        WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, 경로)))
        # print("정상")
    except:
        print("Web wait 함수 에러")
        Slack_Msg("Web wait 함수 에러")
        
df = pd.read_excel("data/전화번호리스트.xlsx")
email = get_secret("email")
password = get_secret("pw")
url = "https://www.modusign.co.kr/"

driver = webdriver.Chrome(executable_path="./chromedriver/chromedriver.exe")
driver.get(url)
time.sleep(2)

# 로그인/
nextXpath('//*[@id="Header"]/div/ul/li[7]/a[1]/b').click()
time.sleep(2)
nextXpath('//*[@id="root"]/div[2]/div/div/div/div/div[1]/form/div[1]/input').send_keys(email)
nextXpath('//*[@id="root"]/div[2]/div/div/div/div/div[1]/form/div[2]/div/input').send_keys(password)
nextXpath('//*[@id="root"]/div[2]/div/div/div/div/div[1]/form/button').send_keys(Keys.ENTER) 
time.sleep(3)

results = []


for i in range(len(df['사번'])):
    try:
        대상자 = df.iloc[i].values
        print(대상자)
        
        # 문서함 초기화면 로딩
        경로 = '//*[@id="root"]/div[2]/div/div[2]/div[2]/section[1]/ul/li[1]/a/div/span'
        web_wait(경로)
        nextXpath(경로).click()
        time.sleep(0.1)

        #템플릿 화면 이동
        경로 = '/html/body/div[1]/div[2]/div/div[2]/div[2]/section[1]/ul/li[3]/a/div/span'   
        web_wait(경로)
        nextXpath(경로).click()
        time.sleep(0.1)
        
        # 첫번째 템플릿 선택
        경로 = '//*[@id="root"]/div[2]/div/div[4]/div/div/div[2]/ul/div[1]/li/ul/li[1]/span'
        web_wait(경로)
        nextXpath(경로).click()

        #서명요청하기 클릭
        경로 = '//*[@id="root"]/div[2]/div/div[4]/div/div/div[2]/ul/div[1]/li/div[2]/div/ul/li[1]'
        web_wait(경로)
        nextXpath(경로).click()
        
        # 확인하고 시작하기 클릭
        num = num_check(int(7+int(i)))
        print(num)
        경로 = f'/html/body/div[{num}]/div/div/div/div/div[2]/button[2]'
        web_wait(경로)
        nextXpath(경로).click()

        # 이름 또는 회사명 입력
        경로 = '//*[@id="first-participant"]/div[1]/div[2]/div/div/div/div/input'
        web_wait(경로)
        nextXpath(경로).click()  
        nextXpath(경로).send_keys(str(대상자[1]))  

        # 카카오톡 전송버튼 클릭
        경로 = '//*[@id="first-participant"]/div[1]/div[3]/button[2]'
        web_wait(경로)
        nextXpath(경로).click()

        # 전화번호 입력(뒤에 여덟자리만)
        경로 = '//*[@id="first-participant"]/div[1]/div[3]/div[2]/div/input'
        web_wait(경로)
        nextXpath(경로).click()
        nextXpath(경로).send_keys(str(대상자[2])) 

        #다음 단계로 버튼 클릭
        경로 = '//*[@id="root"]/div[2]/div/div/div/div/div[3]/div/div/button/span'
        web_wait(경로)
        nextXpath(경로).click()

        #다음 단계로 버튼 클릭
        경로 = '//*[@id="root"]/div[2]/div/div/div/div/div[4]/div/div/button/span'
        web_wait(경로)
        nextXpath(경로).click()
        
        # 입력설정 확인 (서명 없는 전자문서의 경우만 아래 코드 블록 3줄 활성화)
        경로 = '/html/body/div[9]/div/div/div/div/div[2]/button[2]'
        web_wait(경로)
        nextXpath(경로).click()
        
        # 서명유효기간 설정 클릭
        경로 = '//*[@id="root"]/div[2]/div/div/div/div/div[2]/div/div[1]/div/div[2]/p/b'
        web_wait(경로)
        nextXpath(경로).click()

        # 서명유효기간 박스 클릭
        경로 = '//*[@id="root"]/div[2]/div/div/div/div/div[2]/div/div[1]/div/div[2]/div/input'
        web_wait(경로)
        nextXpath(경로).clear()

        # 서명유효기간 입력 (내일까지면 "1", 모레까지면 "2")
        경로 = '//*[@id="root"]/div[2]/div/div/div/div/div[2]/div/div[1]/div/div[2]/div/input'
        web_wait(경로)
        nextXpath(경로).clear()
        nextXpath(경로).send_keys("4")   

        # 설정완료 버튼 클릭
        경로 = '//*[@id="root"]/div[2]/div/div/div/div/div[3]/div/button[2]/span'
        web_wait(경로)
        nextXpath(경로).click()

        # 개인정보수집 동의
        경로 = '//*[@id="agreement-check1"]'
        web_wait(경로)
        nextXpath(경로).click()

        # 서명시 횟수차감 동의
        경로 = '//*[@id="agreement-check2"]'
        web_wait(경로)
        nextXpath(경로).click()

        # 서명요청하기 클릭
        경로 = '/html/body/div[11]/div/div/div/div/div[5]/div/button'
        web_wait(경로)
        nextXpath(경로).click()
        
        time.sleep(1)

        msg = f"{i}번째 {대상자[1]} 발송 완료"
        print(msg)
        results.append(msg)
        
        if (int(i)+1) % 50 == 0:
            Slack_Msg(f"{int(i)+1}번째 발송 완료")
            time.sleep(5)
            
    except:
        print(f"{int(i)+1}번째 {대상자[1]} 발송시 에러 발생")
        Slack_Msg(f"{int(i)+1}번째 {대상자[1]} 발송시 에러 발생")
        break

print("전체 발송 완료")
result_df = pd.DataFrame({"사번": 대상자[0] ,"결과": results})
result_df.to_excel("./results1.xlsx")
Slack_Msg(f"전체 {int(i)+1}명 발송 완료")


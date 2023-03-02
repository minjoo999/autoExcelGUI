import time
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from bs4 import BeautifulSoup
from typing import *

# 절차 전체를 class로 묶음

class autoExcelAdjust():
    def __init__(self, startDate, endDate, title, userId, userPw):
        super().__init__()

        # 각 절차를 클래스 내부의 함수로 정의하고, 그 함수들을 순서대로 __init__에서 실행되도록 하기

        print(startDate, endDate, title, userId, userPw)
        self.all_process(userId, userPw, startDate, endDate)
        # start = self.login(userId, userPw)
        # self.find_purchase(start, startDate, endDate)
    
    # 전체 절차를 한 함수에 담기
    def all_process(self, userId, userPw, startDate, endDate):
        # start = self.login(self, userId, userPw)
        self.login(userId, userPw)
        self.excel_download()
        return "complete"

    # 세부 절차 함수
    # 로그인
    def login(self, userId, userPw):
        global driver
        driver = webdriver.Chrome('chromedriver.exe')

        # 사이트 selenium으로 접속
        driver.get("http://admshop.husstem.co.kr/Login/")

        # 접속한 사이트에서 로그인
        user_id = driver.find_element(By.ID, "txtID")
        user_id.send_keys(userId)

        user_pw = driver.find_element(By.ID, "txtPWD")
        user_pw.send_keys(userPw)

        login_btn = """/html/body/div/form/div/div/button"""
        driver.find_element(By.XPATH, login_btn).click()
        print("로그인 완료")
        return 1


    # 배송완료, 배송중 엑셀출력하기 눌러서 다운로드
    def excel_download(self):

        # 배송완료로 이동
        driver.get("http://admshop.husstem.co.kr/?PG_CODE=140")
        time.sleep(1)

        # 엑셀출력하기 누르기
        excel_btn = """//*[@id="frmSearch"]/div[2]/input"""
        driver.find_element(By.XPATH, excel_btn)

        # 배송완료 끝나면 배송중으로 이동
        driver.get("http://admshop.husstem.co.kr/?PG_CODE=131")
        time.sleep(1)

        # 엑셀출력하기 누르기
        driver.find_element(By.XPATH, excel_btn)

    # 출력된 엑셀파일의 내용을 다른 파일에 적기
    # def edit_final_excel(self):

        # 배송완료로 
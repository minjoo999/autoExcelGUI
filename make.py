import time
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from bs4 import BeautifulSoup
from typing import *

# 절차 전체를 class로 묶음

# startDate = 0
# endDate = 0
# title = ""

class autoExcelAdjust():
    def __init__(self, startDate, endDate, title):
        super().__init__()

        # 각 절차를 클래스 내부의 함수로 정의하고, 그 함수들을 순서대로 __init__에서 실행되도록 하기

        print(startDate, endDate, title)
        start = self.begin()
        self.find_purchase(start, startDate, endDate)

    def begin(self):
        global driver
        driver = webdriver.Chrome('chromedriver.exe')

        # 사이트 selenium으로 접속
        driver.get("http://admshop.husstem.co.kr/Login/")

        # 접속한 사이트에서 로그인
        user_id = driver.find_element(By.ID, "txtID")
        user_id.send_keys("")

        user_pw = driver.find_element(By.ID, "txtPWD")
        user_pw.send_keys("")

        login_btn = """/html/body/div/form/div/div/button"""
        driver.find_element(By.XPATH, login_btn).click()
        print("로그인 완료")
        return 1


    # 페이지 전체 넘기며 날짜 찾기, 원하는 범위 안에 있는 구매 건 찾기 (함수!)
    def find_purchase(self, number, startDate, endDate):
        # time.sleep(1)
        print(startDate, endDate)

        # 배송완료로 이동
        driver.get("http://admshop.husstem.co.kr/?PG_CODE=140")

        # 배송완료 끝나면 배송중으로 이동
        driver.get("http://admshop.husstem.co.kr/?PG_CODE=131")
        time.sleep(1)

        # 구매 건 찾아서 엑셀에 입력하기 (함수!)

        # 페이지 전체 넘기며 날짜 찾기, 원하는 범위 안에 있는 구매 건 찾기, 엑셀에 입력하기 반복

        # 배송완료 마지막 페이지 끝나면 창 닫기

        # 엑셀에서 최종 합산 구하기

        # 파일 저장

        # 날짜로 정산 대상 판별 함수
        first_path = f"""//*[@id="ORDER_{startDate}000000"]/td[1]/div[2]/font/span"""
        

        # //*[@id="ORDER_230224001452"]/td[1]/div[2]
        # //*[@id="ORDER_230224001452"]/td[1]/div[2]/font

    # //*[@id="ORDER_230224001452"]/td[1]/div[2]/font/span  # #ORDER_230224001452 > td:nth-child(1) > div:nth-child(2) > font > span
    # //*[@id="ORDER_230225000754"]/td[1]/div[2]/font/span  # #ORDER_230225000754 > td:nth-child(1) > div:nth-child(2) > font > span

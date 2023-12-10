import time
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from webdriver_manager.chrome import ChromeDriverManager
import os
import getpass
import openpyxl as op
import xlrd

class autoExcelAdjust():
    def __init__(self, startDate, endDate, title, userId, userPw, today):
        super().__init__()

        # 각 절차를 클래스 내부의 함수로 정의하고, 그 함수들을 순서대로 __init__에서 실행되도록 하기

        # print(startDate, endDate, title, userId, userPw, today)
        self.all_process(userId, userPw, startDate, endDate, today, title)

    def all_process(self, userId, userPw, startDate, endDate, today, title):
        self.login(userId, userPw)
        self.excel_download()
        self.edit_final_excel(userId, today, title, startDate, endDate)
        return "complete"

    def login(self, userId, userPw):
        global driver

        service = Service(ChromeDriverManager().install())
        driver = webdriver.Chrome(service=service)

        driver.get("http://admshop.husstem.co.kr/Login/")

        user_id = driver.find_element(By.ID, "txtID")
        user_id.send_keys(userId)

        user_pw = driver.find_element(By.ID, "txtPWD")
        user_pw.send_keys(userPw)

        login_btn = """/html/body/div/form/div/div/button"""
        driver.find_element(By.XPATH, login_btn).click()
        print("로그인 완료")

        # "이미 로그인되어 있습니다" 알림창에 확인버튼 누르기
        try:
            result = driver.switch_to.alert()
            result.accept()
        except:
            pass

    def excel_download(self):

        time.sleep(1)

        # 배송완료로 이동
        driver.get("http://admshop.husstem.co.kr/?PG_CODE=140")
        time.sleep(1)

        # 엑셀출력하기 누르기
        excel_btn = """//*[@id="frmSearch"]/div[2]/input"""
        driver.find_element(By.XPATH, excel_btn).click()

        # 배송완료 끝나면 배송중으로 이동
        driver.get("http://admshop.husstem.co.kr/?PG_CODE=131")
        time.sleep(1)

        # 엑셀출력하기 누르기
        driver.find_element(By.XPATH, excel_btn).click()

    def edit_final_excel(self, userId, today, title, startDate, endDate):

        # 시작일, 종료일 변수 설정
        format = "%Y-%m-%d"
        startDate_2 = time.strptime(startDate, format)
        endDate_2 = time.strptime(endDate, format)

        # 저장경로 (= exe 파일 설치 경로와 동일) 설정
        basePath = os.getcwd()

        # 다운로드 파일 경로 (유저명에 따라 달라짐)
        userName = getpass.getuser()

        downloadPath = "~/Downloads"
        # wb = op.load_workbook(downloadPath + f"/{userId}({today}).xls")
        wb = xlrd.open_workbook(f"~/Downloads/{userId}({today}).xls")
        ws = wb.sheet_by_name("sheet1")

        # 깨진 xls 파일을 어떻게 열어야 하는지 알아보기

        print(ws)
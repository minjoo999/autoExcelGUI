import time
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from bs4 import BeautifulSoup
from typing import *
import win32com.client 

# 절차 전체를 class로 묶음

class autoExcelAdjust():
    def __init__(self, startDate, endDate, title, userId, userPw, today):
        super().__init__()

        # 각 절차를 클래스 내부의 함수로 정의하고, 그 함수들을 순서대로 __init__에서 실행되도록 하기

        print(startDate, endDate, title, userId, userPw, today)
        self.all_process(userId, userPw, startDate, endDate, today)
        # start = self.login(userId, userPw)
        # self.find_purchase(start, startDate, endDate)
    
    # 전체 절차를 한 함수에 담기
    def all_process(self, userId, userPw, startDate, endDate, today):
        # start = self.login(self, userId, userPw)
        # self.login(userId, userPw)
        # self.excel_download()
        self.edit_final_excel(userId, today)
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
    def edit_final_excel(self, userId, today):

        # 양식 열기
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = True
        wb_origin = excel.Workbooks.Open("E:\\2023\\projects\\autoExcelAdjust\\original.xlsx")
        wb_origin_active = wb_origin.ActiveSheet
        # wb_origin_active = wb_origin.WorkSheet("Sheet1")

        # 배송완료 엑셀파일 열기
        ship_ed = excel.Workbooks.Open(f"C:\\Users\\sjctk\\Downloads\\{userId}(2023-03-02)")
        # ship_complete = excel.Workbooks.Open(f"C:\\Users\\sjctk\\Downloads\\{userId}({today}).xls")
        ship_ed_active = ship_ed.ActiveSheet
        # ship_ed_active = ship_ed.WorkSheets(f"{userId}(2023-03-02)")

        # 특정 항목을 특정 칸에 넣기
        self.edit_excel(ship_ed_active, wb_origin_active)

        # 배송중 엑셀파일 열기
        # ship_ing = excel.Workbooks.Open(f"C:\\Users\\sjctk\\Downloads\\{userId}({today}) (1).xls")
        ship_ing = excel.Workbooks.Open(f"C:\\Users\\sjctk\\Downloads\\{userId}(2023-03-02) (1)")
        ship_ing_active = ship_ing.ActiveSheet

        # 특정 항목을 특정 칸에 넣기
        # self.edit_excel(ship_ing_active, wb_origin_active)

        excel.Quit()

    # 특정 항목 -> 특정 칸 절차 세부 분류
    def edit_excel(self, read_file, write_file):

        # 시작점을 어떻게 선정하는가? 첫번째 파일은 그냥 시작점 (AB2, B3 등)을 하드코딩하면 됨
        # 두번째 파일이 문제 -> 첫번째 파일을 집어넣은 그 마지막 행 번호를 변수로 지정하고 불러오면 될 듯.
        # 첫번째 파일일 경우의 로직과 두번째 파일일 경우의 로직을 if로 분리해보기. 아니면 변수명 저장을 잘 건드려보고.

        # len으로 내용이 있는 만큼만...?
        # 사이트 다운로드 파일 파트는 2행으로 시작하는 하드코딩 하기
        # 작성하는 파일 파트는 CurrentRegion 쓰고, 그 마지막 숫자 다음 줄부터 작업 시작시키기

        # Select()
        write_file.UsedRange.Select()
        
        # write_file.Range("B2", "F2").select()
        # print(read_file.UsedRange.SpecialCells(11).value)

        # if write_file.UsedRange.SpecialCells(11).value == None:


        # 주문완료일자: 홈페이지 파일 AB2 -> 정산자료 B3
        # for n in range(2, ):
        #     read_file.Range(f"AB{n}").Value = write_file.Range(f"B{n+1}")
        
        # 판매상품: 홈페이지 파일 O2 -> 정산자료 C3

        # 판매금액: 홈페이지 파일 S2 -> 정산자료 D3

        # 총샵마진: 홈페이지 파일 Y2 -> 정산자료 G3

        # 주문인: 홈페이지 파일 F2 -> 정산자료 J3

        # 수령인: 홈페이지 파일 J2 -> 정산자료 K3

        # return True
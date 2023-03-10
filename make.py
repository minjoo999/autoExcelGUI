import time
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from bs4 import BeautifulSoup
from typing import *
import win32com.client 
from datetime import datetime

# 절차 전체를 class로 묶음

class autoExcelAdjust():
    def __init__(self, startDate, endDate, title, userId, userPw, today):
        super().__init__()

        # 각 절차를 클래스 내부의 함수로 정의하고, 그 함수들을 순서대로 __init__에서 실행되도록 하기

        print(startDate, endDate, title, userId, userPw, today)
        self.all_process(userId, userPw, startDate, endDate, today, title)
    
    # 전체 절차를 한 함수에 담기
    def all_process(self, userId, userPw, startDate, endDate, today, title):
        self.login(userId, userPw)
        self.excel_download()
        self.edit_final_excel(userId, today, title, startDate, endDate)
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

        # "이미 로그인되어 있습니다" 알림창에 확인버튼 누르기
        try:
            result = driver.switch_to.alert()
            result.accept()
        except:
            pass


    # 배송완료, 배송중 엑셀출력하기 눌러서 다운로드
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

    # 출력된 엑셀파일의 내용을 다른 파일에 적기
    # 저장경로 소프트코딩 방법 고민하기
    def edit_final_excel(self, userId, today, title, startDate, endDate):

        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = True
        # excel.Visible = False

        # 시작일, 종료일 변수 설정
        format = "%Y-%m-%d"
        startDate_2 = time.strptime(startDate, format)
        endDate_2 = time.strptime(endDate, format)

        # 양식 열기
        wb_origin = excel.Workbooks.Open("E:\\2023\\projects\\autoExcelAdjust\\original.xlsx")
        wb_origin_active = wb_origin.ActiveSheet

        # 배송완료 엑셀파일 열기
        ship_ed = excel.Workbooks.Open(f"C:\\Users\\sjctk\\Downloads\\{userId}({today}).xls")

        # 배송완료 변수 지정, 특정 항목을 특정 칸에 넣기 시작
        shipped_active = ship_ed.ActiveSheet
        self.edit_excel(shipped_active, wb_origin_active, startDate_2, endDate_2)

        # 배송중 엑셀파일 열기
        ship_ing = excel.Workbooks.Open(f"C:\\Users\\sjctk\\Downloads\\{userId}({today}) (1).xls")
        
        # 배송중 변수 지정, 특정 항목을 특정 칸에 넣기 시작
        shipping_active = ship_ing.ActiveSheet
        self.edit_excel(shipping_active, wb_origin_active, startDate_2, endDate_2)

        # 최종 제작된 파일 저장, 절차 종료
        wb_origin_active.SaveAs(f"E:\\2023\\projects\\autoExcelAdjust\\filesave\\{title}.xlsx")
        excel.Quit()

    # 특정 항목 -> 특정 칸 절차 세부 분류
    def edit_excel(self, read_file, write_file, startDate, endDate):

        # 다운로드된 파일에 구매내역 써있는 줄의 수
        lines = read_file.UsedRange.CurrentRegion.Rows.Count

        global num_final
        num_final = 0

        # 첫번째 파일, 두번째 파일 구분
        if type(write_file.Range("B3").Value) == "NoneType":

            # num을 세서 남겨야 함
            self.copy_and_paste(read_file, write_file, lines, startDate, endDate, 0)
                    
        else:
            # 앞에서 세놓은 num을 줄수에 포함시켜야 함.
            self.copy_and_paste(read_file, write_file, lines, startDate, endDate, num_final)

            # 읽히는 파일 입장에서는 2행, 3행, 4행 이런 식일텐데
            # 쓰는 파일은 이제 앞에 해놓은 결과에 따라 몇행에서 시작할지 정해지는 거지
            # 미리 적어놓은게 몇행까지인가...를 알수 있는 방법은?


    # 파일 복붙 로직 (AB2 -> B3 등)
    def copy_and_paste(self, read_file, write_file, lines, startDate, endDate, line_num):

        # 반복 횟수 누적해서 기록해놓기. 여기서 기록한 숫자를 global 변수 지정해서 다음 번에 row 갯수로 써먹을수 있게.

        # 아래 로직 함수로 만들기
        # 첫번째 파일: AB2 -> B3
        # 이미 첫번째 파일에서 5줄을 복붙해놓은 이후: AB2 -> B8 (AB2 -> B(3+5))
        # line + 1 + 0 vs. line + 1 + (앞에서구한)num

        num = 0
        format = "%Y-%m-%d"

        for line in range(2, lines+1):

            # 정산 대상에 해당하는 날짜 지목
            # 그 날짜에 해당하는 셀들만 복붙하기
            time_read = read_file.Range(f"AB{line}").Value.strftime(format)
            time_read_file = time.strptime(time_read, format)

            # print(read_file.Range("AB2").Select())

            # Select 메소드는 오직! write_file 에서만 오류가 생김
            # time_read를 for문 안에 넣으니까 write_file의 select가 사고나네...
            # print(write_file.Range("B2").Select())

            time.sleep(1)

            if startDate <= time_read_file <= endDate:

                # 주문완료일자: 홈페이지 파일 AB2 -> 정산자료 B3
                read_file.Range(f"AB{line}").Copy()
                write_file.Range(f"B{line + 1 + line_num}").Select()
                write_file.Paste()

                # 판매상품: 홈페이지 파일 O2 -> 정산자료 C3
                read_file.Range(f"O{line}").Copy()
                write_file.Range(f"C{line + 1 + line_num}").Select()
                write_file.Paste()

                # 판매금액: 홈페이지 파일 S2 -> 정산자료 D3
                read_file.Range(f"S{line}").Copy()
                write_file.Range(f"D{line + 1 + line_num}").Select()
                write_file.Paste()

                # 판매수량: 홈페이지 파일 R2 -> 정산자료 E3
                read_file.Range(f"R{line}").Copy()
                write_file.Range(f"E{line + 1 + line_num}").Select()
                write_file.Paste()

                # 총샵마진: 홈페이지 파일 Y2 -> 정산자료 G3
                read_file.Range(f"Y{line}").Copy()
                write_file.Range(f"G{line + 1 + line_num}").Select()
                write_file.Paste()

                # 주문인: 홈페이지 파일 F2 -> 정산자료 J3
                read_file.Range(f"F{line}").Copy()
                write_file.Range(f"J{line + 1 + line_num}").Select()
                write_file.Paste()

                # 수령인: 홈페이지 파일 J2 -> 정산자료 K3
                read_file.Range(f"J{line}").Copy()
                write_file.Range(f"K{line + 1 + line_num}").Select()
                write_file.Paste()

                num = num + 1

        num_final = num
        print(num_final)
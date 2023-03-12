import time
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from bs4 import BeautifulSoup
from typing import *
import win32com.client 
import os
import getpass

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

        global excel
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = True

        # 시작일, 종료일 변수 설정
        format = "%Y-%m-%d"
        startDate_2 = time.strptime(startDate, format)
        endDate_2 = time.strptime(endDate, format)

        # 저장경로 (= exe 파일 설치 경로와 동일) 설정
        basePath = os.getcwd()

        # 다운로드 파일 경로 (유저명에 따라 달라짐)
        userName = getpass.getuser()

        # 배송완료 엑셀파일 열기
        downloadPath = "C:\\Users\\" + userName + "\\Downloads"
        ship_ed = excel.Workbooks.Open(downloadPath + f"\\{userId}({today}).xls")
        shipped_active = ship_ed.ActiveSheet

        # 양식 열고, 입력 시작
        filename = "original.xlsx"
        origin_path = os.path.join(basePath, filename)
        wb_origin = excel.Workbooks.Open(origin_path)
        wb_origin_active = wb_origin.ActiveSheet

        self.edit_excel(shipped_active, wb_origin_active, startDate_2, endDate_2)

        # 저장하고 양식 닫기
        save_path = basePath + f"\\filesave\\editing\\editing_{today}.xlsx"
        wb_origin_active.SaveAs(save_path)
        wb_origin.Close()

        # 배송중 엑셀파일 열기
        ship_ing = excel.Workbooks.Open(downloadPath + f"\\{userId}({today}) (1).xls")
        shipping_active = ship_ing.ActiveSheet
        
        # 아까 편집한 양식 열고, 입력 시작
        wb_origin2 = excel.Workbooks.Open(save_path)
        wb_origin2_active = wb_origin2.ActiveSheet
        self.edit_excel(shipping_active, wb_origin2_active, startDate_2, endDate_2)

        # 정산일자 적기
        wb_origin2_active.Range("C35").Value = f"정산일자: {startDate} ~ {endDate}"

        # 저장하고 절차 종료
        final_save = f"\\filesave\\final\\{title}.xlsx"
        wb_origin2_active.SaveAs(basePath + final_save)
        excel.Quit()

    # 특정 항목 -> 특정 칸 절차 세부 분류
    def edit_excel(self, read_file, write_file, startDate, endDate):

        # 다운로드된 파일에 구매내역 써있는 줄의 수
        lines = read_file.UsedRange.CurrentRegion.Rows.Count

        # 작성할 양식 파일에 내용이 차있는 줄의 수 (정수 변환)
        written_lines = excel.WorkSheetFunction.CountA(write_file.Range("B:B"), 0)
        written_lines = round(written_lines)
        print(written_lines)

        self.copy_and_paste(read_file, write_file, lines, startDate, endDate, written_lines)

    # 파일 복붙 로직 (AB2 -> B3 등)
    def copy_and_paste(self, read_file, write_file, lines, startDate, endDate, line_num):

        num = 0
        format = "%Y-%m-%d"

        for line in range(2, lines+1):

            # 정산 대상에 해당하는 날짜 지목
            # 그 날짜에 해당하는 셀들만 복붙하기
            time_read = read_file.Range(f"AB{line}").Value.strftime(format)
            time_read_file = time.strptime(time_read, format)

            # Select 메소드는 오직! write_file 에서만 오류가 생김 -> write_file을 read_file 위에 펼쳐지게 로직 수정

            if startDate <= time_read_file <= endDate:

                read_line = line + line_num - 1

                # line_num = 2 (양식 원본) -> read_line = 2 + 2 - 1 = 3
                # line_num = 18 (이미 채워놓은 양식) -> read_line = 2 + 18 - 1 = 19

                print(read_line)

                # 주문완료일자: 홈페이지 파일 AB2 -> 정산자료 B3
                read_file.Range(f"AB{line}").Copy()
                write_file.Range(f"B{read_line}").Select()
                write_file.Paste()

                # 판매상품: 홈페이지 파일 O2 -> 정산자료 C3
                read_file.Range(f"O{line}").Copy()
                write_file.Range(f"C{read_line}").Select()
                write_file.Paste()

                # 판매금액: 홈페이지 파일 S2 -> 정산자료 D3
                read_file.Range(f"S{line}").Copy()
                write_file.Range(f"D{read_line}").Select()
                write_file.Paste()

                # 판매수량: 홈페이지 파일 R2 -> 정산자료 E3
                read_file.Range(f"R{line}").Copy()
                write_file.Range(f"E{read_line}").Select()
                write_file.Paste()

                # 총샵마진: 홈페이지 파일 Y2 -> 정산자료 G3
                read_file.Range(f"Y{line}").Copy()
                write_file.Range(f"G{read_line}").Select()
                write_file.Paste()

                # 주문인: 홈페이지 파일 F2 -> 정산자료 J3
                read_file.Range(f"F{line}").Copy()
                write_file.Range(f"J{read_line}").Select()
                write_file.Paste()

                # 수령인: 홈페이지 파일 J2 -> 정산자료 K3
                read_file.Range(f"J{line}").Copy()
                write_file.Range(f"K{read_line}").Select()
                write_file.Paste()

                num = num + 1

        

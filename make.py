import time
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from bs4 import BeautifulSoup
from typing import *

driver = webdriver.Chrome('chromedriver.exe')

# 사이트 selenium으로 접속
driver.get("http://admshop.husstem.co.kr/Login/")

# 접속한 사이트에서 로그인
# user_id = 

# 배송중으로 이동

# 페이지 전체 넘기며 날짜 찾기, 원하는 범위 안에 있는 구매 건 찾기 (함수!)

# 구매 건 찾아서 엑셀에 입력하기 (함수!)

# 배송중 마지막 페이지 끝나면 배송완료로 이동

# 페이지 전체 넘기며 날짜 찾기, 원하는 범위 안에 있는 구매 건 찾기, 엑셀에 입력하기 반복

# 배송완료 마지막 페이지 끝나면 창 닫기

# 엑셀에서 최종 합산 구하기

# 파일 저장


def begin():
    start = "작업 시작합니다"
    return start
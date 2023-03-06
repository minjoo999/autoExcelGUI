import sys
from PyQt5.QtWidgets import *
from PyQt5 import uic
from datetime import datetime
from make import autoExcelAdjust
 
# UI파일 연결
form_class = uic.loadUiType("autoUi.ui")[0]

# 화면을 띄우는데 사용되는 Class 선언
class WindowClass(QMainWindow, form_class) :
    def __init__(self) :
        super().__init__()
        self.setupUi(self)

        # 기본 제목
        self.checkBox.stateChanged.connect(self.defaultTitle)

        # 내용 확정 및 작동 시작 버튼 (제목이 있어야만 버튼 작동)
        self.fixBtn.clicked.connect(self.fixBtnPush)
        self.startBtn.clicked.connect(self.startBtnPush)
    
    # 기본 제목 정하기
    def defaultTitle(self):
        global today
        today = datetime.now().date()
        self.titleText.setPlainText(f"시더스 정산자료_스타제과_{today}")

    # 내용 확정 (시작일자, 종료일자, 제목)
    def fixBtnPush(self):
        # startDate = self.startDate.date().toString("yyyy-MM-dd")
        global startDate, endDate, title, userId, userPw
        startDate = self.startDate.date().toString("yyyy-MM-dd")
        endDate = self.endDate.date().toString("yyyy-MM-dd")
        title = self.titleText.toPlainText()
        userId = self.userId.toPlainText()
        userPw = self.userPw.toPlainText()

        # 제목을 지어야만 내용확정 가능
        # 내용 확정 버튼 누르면 내용 변경 막힘
        if len(title) > 0:
            self.startDate.setEnabled(False)
            self.endDate.setEnabled(False)
            self.titleText.setEnabled(False)
            self.userId.setEnabled(False)
            self.userPw.setEnabled(False)
            print(today, startDate, endDate, title, userId, userPw)
        else:
            QMessageBox.warning(self, "경고", "제목을 입력해주세요")

    # 정리 시작
    def startBtnPush(self):

        # 제목을 지어야만 정리 시작 가능
        if len(title) > 0:
            # print(begin())
            # print("작업 시작합니다")
            start = autoExcelAdjust(startDate, endDate, title, userId, userPw, today)
            start
        else:
            QMessageBox.warning(self, "경고", "제목을 입력해주세요")

    # 작업 종료 후 완료했습니다 창 만들기
    # or 만들어진 파일 저장된 폴더 띄우기                                                                                                                  
    
if __name__ == "__main__" :                                                                                                                                                                                  
    app = QApplication(sys.argv) 
    myWindow = WindowClass() 
    myWindow.show()
    app.exec_()
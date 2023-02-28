import sys
from PyQt5.QtWidgets import *
from PyQt5 import uic
from datetime import datetime
import make
 
# UI파일 연결
form_class = uic.loadUiType("autoUi.ui")[0]

# 화면을 띄우는데 사용되는 Class 선언
class WindowClass(QMainWindow, form_class) :
    def __init__(self) :
        super().__init__()
        self.setupUi(self)

        # 기본 제목
        self.checkBox.stateChanged.connect(self.defaultTitle)

        # 내용 확정 및 작동 시작 버튼
        self.fixBtn.clicked.connect(self.fixBtnPush)
        self.startBtn.clicked.connect(self.startBtnPush)
    
    # 기본 제목
    def defaultTitle(self):
        # 제목 정하기
        date = datetime.now().date()
        self.titleText.setPlainText(f"시더스 정산자료_스타제과_{date}")

        # 제목짓기 칸 막기
        self.titleText.setReadOnly(True)
    

    # 내용 확정 (시작일자, 종료일자, 제목)
    def fixBtnPush(self):
        startDate = self.startDate.date()
        endDate = self.endDate.date()
        title = self.titleText.toPlainText()
        print(startDate, endDate, title)

    # 정리 시작
    def startBtnPush(self):
        # start = make.begin
        print(make.begin())

    # 작업 종료 후 완료했습니다 창 만들기
    # or 만들어진 파일 저장된 폴더 띄우기                                                                                                                  
    
if __name__ == "__main__" :                                                                                                                                                                                  
    app = QApplication(sys.argv) 
    myWindow = WindowClass() 
    myWindow.show()
    app.exec_()
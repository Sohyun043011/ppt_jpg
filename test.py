from PyQt5 import uic,QtGui
import os
import sys
import time
from PyQt5.QtWidgets import *
from pptx import Presentation
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor
from pptx.util import Pt
from comtypes import client
from PyQt5.QtWidgets import QWidget, QApplication
from PyQt5.QtGui import QPainter, QPen, QColor, QBrush, QPixmap , QIcon
from PyQt5.QtCore import Qt
import socket
import subprocess
import threading

form_class = uic.loadUiType("ppt_to_jpg.ui")[0] # ppt_to_jpg.ui(xml 형식)에서 레이아웃 및 텍스트 설정값 조정

class MyWindow(QMainWindow, form_class):
    def __init__(self):
        super().__init__()
        self.setupUi(self)
        self.initSetting()
        self.createBtn.clicked.connect(self.createBtn_clicked)
        self.setFixedSize(1600, 850) # 창 사이즈 고정
        self.setWindowTitle('화상회의실 관리 프로그램') # 프로그램 Title 설정
        self.setWindowIcon(QIcon('./wrench.png')) # 프로그램 아이콘 설정
        
    def initSetting(self):
        # 서식 이미지 파일
        pixmap1 = QPixmap("./양식1.jpg")
        pixmap2 = QPixmap("./양식2.jpg")
        pixmap3 = QPixmap("./양식3.jpg")

        self.image1.setPixmap(QPixmap(pixmap1))
        self.image2.setPixmap(QPixmap(pixmap2))
        self.image3.setPixmap(QPixmap(pixmap3))
        
        self.statusLabel.setText('상태표시합니다~')
        self.statusLabel.setStyleSheet("Color : blue")
        
        self.radioBtn_1.toggled.connect(self.onClicked)
        self.radioBtn_2.toggled.connect(self.onClicked)
        self.radioBtn_3.toggled.connect(self.onClicked)
        self.radioBtn_4.toggled.connect(self.onClicked)
        self.radioBtn_1.setChecked(True)
        
    def onClicked(self):
        global pptx_fpath
        global ex_flag
        radioBtn = self.sender()
        if radioBtn.text() == '서식1':
            pptx_fpath = os.path.dirname(os.path.abspath('양식1.pptx'))+'\\양식1.pptx'
            ex_flag = '양식1'
            print('양식1 select' + pptx_fpath)
        elif radioBtn.text() == '서식2':
            pptx_fpath = os.path.dirname(os.path.abspath('양식2.pptx'))+'\\양식2.pptx'
            ex_flag = '양식2'
            print('양식2 select'+pptx_fpath)
        elif radioBtn.text() == '서식3':
            pptx_fpath =  os.path.dirname(os.path.abspath('양식3.pptx'))+'\\양식3.pptx'
            ex_flag = '양식3'
            print('양식3 select'+pptx_fpath)
        else:
            # pptx_fpath =  os.path.dirname(os.path.abspath('양식4.pptx'))+'\\양식4.pptx'
            ex_flag = '양식4'
            print('양식4 select')
            
    def paintEvent(self,event):
        qp = QPainter()
        qp.begin(self)
        #그리기 함수의 호출부분
        self.drawRectangles(qp)
        qp.end()
        
        
    def drawRectangles(self,qp):
        qp.setBrush(QColor(255, 136, 79))
        qp.setPen(QPen(QColor(255, 136, 79), 3))
        qp.drawRect(500, 230, 221, 41)
        qp.drawRect(500,271,41,421)
        qp.drawRect(680,270,41,421)
        
    def createBtn_clicked(self):
        # create 버튼 클릭시 이벤트
        # /팀이름/subject/ 로 폴더 생성
        # label에 넣은 대로 ppt 생성
        subject = self.subject.text()
        self.subject.clear()
        deptLabel = self.deptName.currentText()
        QMessageBox.about(self,"message","/"+deptLabel+"/"+subject)
        

if __name__ == "__main__":
    app = QApplication(sys.argv)
    myWindow = MyWindow()
    myWindow.show()
    app.exec_()
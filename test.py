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
from PyQt5.QtCore import Qt, QTimer
import socket
import subprocess
import threading
import urllib.request
from urllib.parse import urljoin
import webbrowser

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
        self.connect_count=0 # 연결 시도 횟수 설정
        
        self.chrome_path="C:/Program Files (x86)/Google/Chrome/Application/chrome.exe %s" # Chrome 설치 위치
        
        self.timer = QTimer(self)
        self.timer.setInterval(6000)
        self.timer.timeout.connect(self.update_network)
        self.timer.start()
        
        
        
    def initSetting(self):
        # 서식 이미지 파일
        pixmap1 = QPixmap("./양식1.jpg")
        pixmap2 = QPixmap("./양식2.jpg")
        pixmap3 = QPixmap("./양식3.jpg")

        self.image1.setPixmap(QPixmap(pixmap1))
        self.image2.setPixmap(QPixmap(pixmap2))
        self.image3.setPixmap(QPixmap(pixmap3))
        
        self.statusLabel.setText('연결 없음')
        self.statusLabel.setStyleSheet("Color : Red")
        
        # 서식 버튼 설정
        self.radioBtn_1.toggled.connect(self.onClicked)
        self.radioBtn_2.toggled.connect(self.onClicked)
        self.radioBtn_3.toggled.connect(self.onClicked)
        self.radioBtn_4.toggled.connect(self.onClicked)
        self.radioBtn_1.setChecked(True)
        
        self.createBtn.clicked.connect(self.createBtn_clicked)
        
        # 활성화 버튼 설정
        self.wallLinkBtn.clicked.connect(self.onWallOpenClick)
        self.wallActivBtn.clicked.connect(self.onWallActivClick)
        self.nameLinkBtn.clicked.connect(self.onNameOpenClick)
        self.nameActivBtn.clicked.connect(self.onNameActivClick)
        
        
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
        
    def makePPT(self,directory,pptx_fpath,inputValue):
        prs = Presentation(pptx_fpath)                  # 양식 선택시 불러옴
        #슬라이드 1~13 까지 돌면서, inputValue로 부터 받아온 값들을 제목, 부제목에 넣어줌
        #placeholder : Title, Center Title, Subtitle, Body etc
    
        for i in range(13):
            print('--------')
            slide = prs.slides[i]
            inputVal = inputValue[i+1]
            shapes = slide.shapes
            self.text_on_shape(shapes,inputVal)
        prs.save('result.pptx')
    
    def text_on_shape(self,shapes,inputVal):
        # shapes : 한 슬라이드 안
        for shape in shapes:
            if shape.name=="name" or shape.name=="pos":
                # shape 내 텍스트 프레임 선택하기 & 기존 값 삭제하기
                text_frame = shape.text_frame
                p = text_frame.paragraphs[0]
                font_size = p.runs[0].font.size
                # font_color = p.runs[0].font.color.type
                text_frame.clear()
                # 정렬 설정 : 중간정렬
                p.alighnment = PP_ALIGN.CENTER   
                run = p.add_run()
                run.text = inputVal[0] if shape.name=="name" else inputVal[1]
                font = run.font
                font.size = font_size
                # font.color.type = font_color
                
                
    def createBtn_clicked(self):
        # create 버튼 클릭시 이벤트
        # /팀이름/subject/ 로 폴더 생성
        subject = self.subject.text()                           # 폴더 이름
        deptLabel = self.deptName.currentText()                 # 부서명
        directory = os.getcwd()+"\\"+deptLabel+"\\"+subject     # 디렉토리 경로
        inputValue = self.inputValue()                          # 입력값 받아옴
        
    def ping(self, ip):
        if self.connect_count>=6: #시도 횟수 6번 이상이면 0번으로 갱신 후 stop
            self.timer.stop()
            self.connect_count=0
            return False
        self.connect_count+=1
        try:
            print('다음으로 연결 중: http://'+ip)
            urllib.request.urlopen('http://'+ip, timeout=1)
            return True
        except urllib.request.URLError as err:
            return False
    
    def update_network(self): # 스마트월 ip와 스마트명패 ip 각각의 연결성을 확인 후 상태 표시
        wall_ip="192.168.0.60" #스마트월 ip
        name_ip="192.168.0.103/Qname/empMain.aspx?readImage=ok" #스마트명패 ip
        self.nameLinkBtn.setDisabled(True)
        self.wallLinkBtn.setDisabled(True)
        
        if self.ping(wall_ip):
            self.statusLabel.setText('스마트월 연결 성공')
            self.statusLabel.setStyleSheet("Color : Blue")
            self.wallLinkBtn.setDisabled(False)
            self.timer.stop()
            self.connect_count=0
        elif self.ping(name_ip):
            self.statusLabel.setText('스마트명패 연결 성공')
            self.statusLabel.setStyleSheet("Color : Blue")
            self.nameLinkBtn.setDisabled(False)
            self.timer.stop()
            self.connect_count=0
        else:
            self.statusLabel.setText('연결 없음')
            self.statusLabel.setStyleSheet("Color : Red")
            
        
    def onNameActivClick(self): # 스마트명패 활성화 버튼 눌렀을 때 onclick function
        os.startfile('enable.bat.lnk')
        self.timer.start()
        
    def onWallActivClick(self): # 스마트월 활성화 버튼 눌렀을 때 onclick function
        os.startfile('disable.bat.lnk')
        self.timer.start()
    
    def onWallOpenClick(self):
        webbrowser.get(self.chrome_path).open("192.168.0.60")
    
    def onNameOpenClick(self):
        webbrowser.get(self.chrome_path).open("http://192.168.0.103/Qname/empMain.aspx?readImage=ok")
        
if __name__ == "__main__":
    app = QApplication(sys.argv)
    myWindow = MyWindow()
    myWindow.show()
    app.exec_()
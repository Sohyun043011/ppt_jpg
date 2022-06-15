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
from PyQt5.QtGui import QPainter, QPen, QColor, QBrush, QPixmap
from PyQt5.QtCore import Qt
import socket
import subprocess
import threading
from pptx.enum.text import PP_ALIGN   # 정렬 설정하기
from pptx.util import Pt      # Pt 폰트사이즈

form_class = uic.loadUiType("ppt_to_jpg.ui")[0]

class MyWindow(QMainWindow, form_class):
    def __init__(self):
        super().__init__()
        self.setupUi(self)
        self.initSetting()
        
        
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
        
        self.createBtn.clicked.connect(self.createBtn_clicked)
        
        
        
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
        
        if not os.path.exists(directory):
            os.makedirs(directory)
            # 폴더 생성 후, ppt 생성
            self.makePPT(directory,pptx_fpath,inputValue)       #디렉토리 경로,선택한 양식경로, 입력값 
        else: 
            # 이미 있는 폴더인 경우, 이름 다시 설정.
            QMessageBox.about(self,"message",subject+"는 이미 있는 폴더입니다. 다른 이름을 설정해주세요.")
        # label에 넣은 대로 ppt 생성
       
       
        self.subject.clear()
        
    
    # position,name 입력 값 받아오기 inputValue={1:['이름1','직위1'],2:[],...}
    def inputValue(self):
        inputValue= {}
        for i in range(1,14):
            nameChild = self.findChild(QLineEdit,"InputName_%d" % (i)).text()
            namePos = self.findChild(QLineEdit,"InputPos_%d" % (i)).text()
            inputValue[i]=[nameChild,namePos]
        return inputValue
    
        
         
        
        
         

if __name__ == "__main__":
    app = QApplication(sys.argv)
    myWindow = MyWindow()
    myWindow.show()
    app.exec_()
from PyQt5 import uic
import os, shutil
import sys
from PyQt5.QtWidgets import *
from pptx import Presentation
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor
from pptx.util import Pt
from PyQt5.QtGui import QPainter, QPen, QColor, QPixmap , QIcon
from PyQt5.QtCore import QTimer
import urllib.request
import webbrowser
from comtypes import client
import shutil

dataImage_default_path="C:\\Server\\Gachi\\Qname\\dataImage"

form_class = uic.loadUiType("ppt_to_jpg.ui")[0] # ppt_to_jpg.ui(xml 형식)에서 레이아웃 및 텍스트 설정값 조정

class MyWindow(QMainWindow, form_class):
    def __init__(self):
        super().__init__()
        self.setupUi(self)
        self.initSetting()
        self.setFixedSize(1600, 850) # 창 사이즈 고정
        self.setWindowTitle('화상회의실 관리 프로그램') # 프로그램 Title 설정
        self.setWindowIcon(QIcon('./wrench.png')) # 프로그램 아이콘 설정
        self.connect_count=0 # 연결 시도 횟수 설정
        
        self.nameLinkBtn.setDisabled(True) # 초기 링크 버튼 비활성화
        self.wallLinkBtn.setDisabled(True) # 초기 링크 버튼 비활성화
        
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
        self.statusLabel.setStyleSheet("color : red")
        
        # 서식 버튼 설정
        self.radioBtn_1.toggled.connect(self.onClicked)
        self.radioBtn_2.toggled.connect(self.onClicked)
        self.radioBtn_3.toggled.connect(self.onClicked)
        # self.radioBtn_1.setChecked(True)
        self.selectForm.clicked.connect(self.onClickSelect)
        
        self.createBtn.clicked.connect(self.createBtn_clicked)
        self.deleteBtn.clicked.connect(self.deleteBtn_clicked)
        
        # 활성화 버튼 설정
        self.wallLinkBtn.clicked.connect(self.onWallOpenClick)
        self.wallActivBtn.clicked.connect(self.onWallActivClick)
        self.nameLinkBtn.clicked.connect(self.onNameOpenClick)
        self.nameActivBtn.clicked.connect(self.onNameActivClick)
        
        self.set_style()
    
    def set_style(self):
        with open("update_style", 'r') as f:
            self.setStyleSheet(f.read())    
        
    def onClicked(self):
        global pptx_fpath
        global ex_flag
        radioBtn = self.sender()
        if radioBtn.text() == '서식1':
            pptx_fpath = os.path.dirname(os.path.abspath('양식1.pptx'))+'\\양식1.pptx'
            ex_flag = '양식1'
            # print('양식1 select' + pptx_fpath)
        elif radioBtn.text() == '서식2':
            pptx_fpath = os.path.dirname(os.path.abspath('양식2.pptx'))+'\\양식2.pptx'
            ex_flag = '양식2'
            # print('양식2 select'+pptx_fpath)
        elif radioBtn.text() == '서식3':
            pptx_fpath =  os.path.dirname(os.path.abspath('양식3.pptx'))+'\\양식3.pptx'
            ex_flag = '양식3'
            # print('양식3 select'+pptx_fpath)
        else:
            # pptx_fpath =  os.path.dirname(os.path.abspath('양식4.pptx'))+'\\양식4.pptx'
            ex_flag = '그 외'
            print('그 외 select')
            
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
        
    def makePPT(self,directory,subject,pptx_fpath,inputValue):
        prs = Presentation(pptx_fpath)                  # 양식 선택시 불러옴
        #슬라이드 1~15 까지 돌면서, inputValue로 부터 받아온 값들을 제목, 부제목에 넣어줌(14,15는 빈 슬라이드)
        #placeholder : Title, Center Title, Subtitle, Body etc
    

        for i in range(15):
            slide = prs.slides[i]
            inputVal = inputValue[i+1]
            shapes = slide.shapes
            self.text_on_shape(shapes,inputVal)
        prs.save(directory+'\\'+subject+'.pptx')
        
    def makeJPG(self,directory,subject):
        ppt = client.CreateObject('Powerpoint.Application')
        ppt.Presentations.Open(directory+"\\"+subject+".pptx")
        ppt.ActivePresentation.Export(directory, 'JPG')
        ppt.ActivePresentation.Close()
        ppt.Quit()
        
        # 슬라이드1~15까지 존재.->1.jpg, 2.jpg,...15.jpg 이거를 복사
        for i in range(15):
            os.rename(directory+"\\슬라이드" + str(i+1) + ".JPG", directory + "\\" + str((i+1)) + '.jpg')
            shutil.copyfile(directory+"\\"+str((i+1)) + '.jpg',directory+"\\"+str(15+(i+1)) + '.jpg')
        
    def inputValue(self):
        inputValue= {}
        for i in range(1,14):
            nameChild = self.findChild(QLineEdit,"InputName_%d" % (i)).text()
            namePos = self.findChild(QLineEdit,"InputPos_%d" % (i)).text()
            inputValue[i]=[nameChild,namePos]
        inputValue[14]=['','']
        inputValue[15]=['','']
        return inputValue
    
    def onClickSelect(self):
        # QMessageBox.about(self,"message",'select vox')
        # 서식 찾기 -> 파일 탐색기 열기-> 파일 선택-> 그 파일 경로 : pptx_fpath로 설정
        # select_folder = QFileDialog.getExistingDirectory(self, '폴더를 선택해주세요', dataImage_default_path, QFileDialog.ShowDirsOnly)
        global pptx_fpath
        select_file = QFileDialog.getOpenFileName(self) 
        print(select_file[0])
        pptx_fpath = select_file[0]
        if pptx_fpath=='':
            QMessageBox.about(self,"message","파일이 선택되지 않았습니다.다시 선택해주세요.")
        else: self.findFormLabel.setText("선택 서식 : \n"+pptx_fpath)
    
    def text_on_shape(self,shapes,inputVal):
        # shapes : 한 슬라이드 안
        for shape in shapes:
            if shape.name=="name" or shape.name=="pos":
                # shape 내 텍스트 프레임 선택하기 & 기존 값 삭제하기
                text_frame = shape.text_frame
                p = text_frame.paragraphs[0]
                font_size = p.runs[0].font.size.pt
                font_color = p.runs[0].font.color
                font_bold = p.runs[0].font.bold
                font_name = p.runs[0].font.name
                # font_brightness =  p.runs[0].font.brightness
                # print(font_color)
                # print(font_color.brightness)
                text_frame.clear()
                # 정렬 설정 : 중간정렬
                p.alighnment = PP_ALIGN.CENTER   
                run = p.add_run()
                run.text = inputVal[0] if shape.name=="name" else inputVal[1]
                font = run.font 
                font.name = font_name
                
                font.size = Pt(font_size)
                if font_bold==True:
                    # bold 설정 되어있다면
                    font.bold = font_bold
                if font_color.type!=None:
                    # 블랙아닌 경우
                    # SCHEME인 경우
                    
                    if font_color.theme_color!=0:
                        font.color.theme_color=font_color.theme_color
                        font.color.brightness = font_color.brightness
                        
                    else:
                        # RGB인 경우
                        # print(int(f'{font_color.rgb}',16))
                        # print(int(str(font_color.rgb)[0:2],16))
                        
                        font.color.rgb = RGBColor(int(str(font_color.rgb)[0:2],16),int(str(font_color.rgb)[2:4],16),int(str(font_color.rgb)[4:6],16))
                        font.color.brightness = font_color.brightness
                else:
                    # 블랙인 경우,
                    font.color.rgb = RGBColor(0,0,0)
                    
    def disableBtn(func):
        def wrapper(self):
            self.nameLinkBtn.setDisabled(True) # 초기 링크 버튼 비활성화
            self.wallLinkBtn.setDisabled(True) # 초기 링크 버튼 비활성화
            func(self)
        return wrapper            
                
    def createBtn_clicked(self):
        # create 버튼 클릭시 이벤트
        # /팀이름/subject/ 로 폴더 생성
        subject = self.subject.text()                           # 폴더 이름
        if subject=='':
            # 공백인 경우 입력하게 하기
            QMessageBox.about(self,"message","폴더명을 입력해주세요.")
        else:
            deptLabel = self.deptName.currentText()                 # 부서명
            # directory = dataImage_default_path+"\\"+deptLabel+"\\"+subject     # 디렉토리 경로
            directory = os.getcwd()+"\\"+deptLabel+"\\"+subject
            inputValue = self.inputValue()     
            if not os.path.exists(directory):
                os.makedirs(directory)
                # 폴더 생성 후, ppt 생성
                self.makePPT(directory,subject,pptx_fpath,inputValue)       #디렉토리 경로,선택한 양식경로, 입력값
                self.makeJPG(directory,subject) 
            else: 
                # 이미 있는 폴더인 경우, 이름 다시 설정.
                QMessageBox.about(self,"message",subject+"는 이미 있는 폴더입니다. 다른 이름을 설정해주세요.")
            self.subject.clear()
        
    def deleteBtn_clicked(self):
        # create 버튼 클릭시 이벤트
        folderPath = QFileDialog.getExistingDirectory(self, '폴더를 선택해주세요', dataImage_default_path, QFileDialog.ShowDirsOnly)
        print(folderPath.split('/')[-2])
        
        
        if os.path.exists(folderPath): # 폴더가 존재할 때
            # 폴더 상위폴더가 dataImage일 경우(부서 카테고리를 지우려고 하는 경우)
            if folderPath.split('/')[-2]=='dataImage':
                QMessageBox.warning(self, '알림', '부서 카테고리는 삭제할 수 없습니다. 다시 시도해주세요.')
                return
            if '/'.join(folderPath.split('/')[:-2])!='C:/Server/Gachi/Qname/dataImage':
                QMessageBox.warning(self, '알림', '다른 디렉토리에 존재하는 파일은 삭제할 수 없습니다. 다시 시도해주세요.')
                return
            
            buttonReply=QMessageBox.information(self,'알림','정말로 해당 파일을 삭제하시겠습니까?', QMessageBox.Yes|QMessageBox.No, QMessageBox.No)
            if buttonReply==QMessageBox.Yes:
                shutil.rmtree(folderPath) # 폴더 하위에 파일의 유무에 관계없이 무조건 삭제
                QMessageBox.information(self,'알림','폴더와 하위 파일들이 삭제되었습니다.')
        elif folderPath=='':
            return
        else:
            QMessageBox.information(self,'알림','폴더가 존재하지 않습니다.')
        
    def ping(self, ip):
        if self.connect_count>=4: #시도 횟수 2번 이상이면 0번으로 갱신 후 stop (한번 시도마다 2개의 ip에 대해 조사)
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
            self.statusLabel.setStyleSheet("Ccolor : blue;")
            self.wallLinkBtn.setDisabled(False)
            self.timer.stop()
            self.connect_count=0
        elif self.ping(name_ip):
            self.statusLabel.setText('스마트명패 연결 성공')
            self.statusLabel.setStyleSheet("color : blue;")
            self.nameLinkBtn.setDisabled(False)
            self.timer.stop()
            self.connect_count=0
        else:
            self.statusLabel.setText('연결 없음')
            self.statusLabel.setStyleSheet("color : red;")
            
    @disableBtn        
    def onNameActivClick(self): # 스마트명패 활성화 버튼 눌렀을 때 onclick function
        QMessageBox.information(self,'알림','스마트명패가 활성화되었습니다. 아래 링크 버튼이 활성화가 될 때까지 잠시만 기다려주세요.')
        os.startfile('enable.bat.lnk')
        self.timer.start()

    @disableBtn
    def onWallActivClick(self): # 스마트월 활성화 버튼 눌렀을 때 onclick function
        os.startfile('disable.bat.lnk')
        QMessageBox.information(self,'알림','스마트월이 활성화되었습니다. 아래 링크 버튼이 활성화가 될 때까지 잠시만 기다려주세요.')
        self.timer.start()
    
    def onWallOpenClick(self):
        webbrowser.get(self.chrome_path).open("192.168.0.60")
    
    def onNameOpenClick(self):
        webbrowser.get(self.chrome_path).open("http://192.168.0.103/Qname/empMain.aspx?readImage=ok")
    
    
    
if __name__ == "__main__":
    try:
        app = QApplication(sys.argv)
        myWindow = MyWindow()
        myWindow.show()
        app.exec_()
    except Exception as e:
        print(e)
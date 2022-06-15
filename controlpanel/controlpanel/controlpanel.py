import tkinter
import os
import webbrowser
from functools import partial
import socket
import urllib.request

class Window():
    def __init__(self):
        self.window=tkinter.Tk()
        self.chrome_path="C:/Program Files (x86)/Google/Chrome/Application/chrome.exe %s"
        self._response_flag=0 #0: 연결 x, 1: 스마트월 연결, 2: 스마트명패 연결
        self.response=0

        self.window_configure()
        self.label_configure()

    def window_configure(self):
        self.window.configure(bg="white")
        self.window.title("디지털회의실 관리자 지원 프로그램")
        self.window.geometry("390x300+750+350")
        self.window.resizable(False, False)

    def label_configure(self):
        self.wall = tkinter.PhotoImage(file = "logo.gif")
        self.imagelabel=tkinter.Label(self.window, image=self.wall, borderwidth=0)
        self.imagelabel.place(x=10,y=10)
        self.statuslabel=tkinter.Label(self.window, text="", background="white", borderwidth=1, fg="blue")
        self.statuslabel.place(x=240,y=10, width=120, height=45)
        self.statuslabel.config(anchor="center")

        self.b1=tkinter.Button(self.window, text="스마트월 활성화", command=self.onClick_b1)
        self.b2=tkinter.Button(self.window, text="스마트명패 활성화", command=self.onClick_b2)
        self.b3=tkinter.Button(self.window, text="스마트월 링크", command=self.onClick_b3)
        self.b4=tkinter.Button(self.window, text="스마트명패 링크", command=self.onClick_b4)

        self.b1.place(x=10, y=70, width=180, height=50)
        self.b2.place(x=200, y=70, width=180, height=50)
        self.b3.place(x=10, y=130, width=180, height=50)
        self.b4.place(x=200, y=130, width=180, height=50)

        self.text=tkinter.Text(self.window, width=52, height=7)
        self.text.pack(side='bottom', pady=10)
        self.text.insert(tkinter.END, "스마트월 또는 스마트명패를 이용하기 전 반드시 활성화버튼을 눌러주시기 바랍니다.\
        \n스마트명패를 사용하기 위해서는 랜케이블을 연결해주셔야 합니다.\
        \n활성화 버튼을 누른 후에도 접속이 되지 않는다면 한번 더 활성화 버튼을 눌러주시거나\
프로그램을 종료 후 다시실행해주시기 바랍니다.")
    
    def update_network(self):
        
        self._response_flag=0
        wall_ip="192.168.0.60" #스마트월 ip
        name_ip="192.168.0.103/Qname/empMain.aspx?readImage=ok" #스마트명패 ip
        
        self.response = self.ping(wall_ip)
        if self.response is True:
            self._response_flag = 1
            self.statuslabel.configure(text="스마트월 연결 성공")
            self.window.after(6000, self.update_network)
            return
    
        self.response = self.ping(name_ip)
        if self.response is True:
            self._response_flag = 2
            self.statuslabel.configure(text="스마트명패 연결 성공")
        else:
            self.statuslabel.configure(text="연결 없음")
        self.window.after(6000, self.update_network)
        return
        
        
        
    def ping(self, ip):
        try:
            urllib.request.urlopen('http://'+ip, timeout=1)
            return True
        except urllib.request.URLError as err:
            return False

    def onClick_b1(self):
        os.startfile('disable.bat.lnk')

    def onClick_b2(self):
        os.startfile('enable.bat.lnk')

    def onClick_b3(self):
        # print("b2 작동")
        url="192.168.0.60"
        webbrowser.get(self.chrome_path).open(url)

    def onClick_b4(self):   
        # print("b2 작동")
        url="http://192.168.0.103/Qname/empMain.aspx?readImage=ok"
        webbrowser.get(self.chrome_path).open(url)

app=Window()
app.update_network()
app.window.mainloop()  

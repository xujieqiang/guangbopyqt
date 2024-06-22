import sys
import time

import wm
from PyQt5.QtWidgets import QApplication, QMainWindow, QLabel
from pydub import AudioSegment
from PyQt5.QtWidgets import *
import xlwt



starttime=''
lasttime=0
name=''
medias=''
data=[]
onerec=[]
count_num=0
tag=0
ntt=''
dict={
    '1;':'初一',
    '2;':'初二',
    '3;':'初三',
    '4;':'全校',
    '5;':'监听室',
    '7;':'行政楼、科技楼',
    '8;':'室外广播'
}
play_area=[]
def getname():
    global name
    name=ui.lineEdit.text()
# def getcombobox_value():
#     global groups
#     txt=ui.comboBox.currentText()
#     groups=txt
#     print(groups)
#     return groups
#     #ui.lineEdit.setText(groups)

def getstarttime():
    global starttime
    starttime=ui.timeEdit.text()

def getplay_area():
    global play_area
    global tag
    if tag==0:
        if ui.checkBox.isChecked():
            play_area.append(ui.checkBox.text())
        if ui.checkBox_2.isChecked():
            play_area.append(ui.checkBox_2.text())
        if ui.checkBox_3.isChecked():
            play_area.append(ui.checkBox_3.text())
        if ui.checkBox_4.isChecked():
            play_area.append(ui.checkBox_4.text())
        if ui.checkBox_5.isChecked():
            play_area.append(ui.checkBox_5.text())
        if ui.checkBox_6.isChecked():
            play_area.append(ui.checkBox_6.text())
        if ui.checkBox_7.isChecked():
            play_area.append(ui.checkBox_7.text())
        tag=1

def getallvalues():
    global medias
    getname()
    getstarttime()
    medias=ui.lineEdit_2.text()
    getplay_area()


def settablevalue(endtime):
    global count_num
    global ntt
    nameitem=QTableWidgetItem(name)
    i=count_num
    ui.tableWidget.setItem(i,0,nameitem)
    timeitem=QTableWidgetItem(starttime)
    ui.tableWidget.setItem(i,1,timeitem)
    mediaitem= QTableWidgetItem(medias)
    ui.tableWidget.setItem(i, 2, mediaitem)
    enditem = QTableWidgetItem(endtime[11:])
    ui.tableWidget.setItem(i, 3, enditem)

    playarea_to_text()
    gitem=QTableWidgetItem(ntt)
    ui.tableWidget.setItem(i,4,gitem)
    count_num+=1

def readfile():
    global lasttime
    global medias
    fileName_choose, filetype = QFileDialog.getOpenFileName(
                                                            None,
                                                            'D:\\' , # 起始路径
                                                            'D:\\',
                                                            "All Files (*);;Mp3 Files (*.mp3)")  # 设置文件扩展名过滤,用双分号间隔
    l=0.0
    if fileName_choose == "":
        print("\n取消选择")
        return
    else:
        medias=fileName_choose
        ui.lineEdit_2.setText(fileName_choose)

    l=get_duration_mp3(medias)
    lasttime=int(l)
    print("\n你选择的文件为:")
    print(fileName_choose)


    print("文件筛选器类型: ", filetype)


# 获取MP3文件的时间长度
def get_duration_mp3(audio_file):
    audio = AudioSegment.from_file(audio_file)
    duration_milliseconds = len(audio)
    duration_seconds = duration_milliseconds / 1000.0
    return duration_seconds


def playarea_to_text():
    global ntt
    global play_area
    for v in play_area:
        ntt = ntt + str(v)+' -> '

## 将字符串的starttime  加上 MP3的时间长度，最终获得 音乐终止时间
def changestrtotime(stime):
    global lasttime
    nstr='2023-10-10 '+stime
    dt1 = time.strptime(nstr, "%Y-%m-%d %H:%M:%S")
    timeint=time.mktime(dt1)
    newtimeint=timeint+lasttime
    t=time.localtime(newtimeint)
    dt = time.strftime("%Y-%m-%d %H:%M:%S", t)
    return dt

def area_to_num():
    global play_area
    txt=''
    for v in play_area:
        for k,val in dict.items():
            if val==v:
                txt+=k
    print(txt)
    return txt
def generate_data():
    ## columns = ['Name', 'JobType', 'JobMask', 'Duration', 'StartTime', 'StopTime',
    ##           'JobData', 'RepeatNum', 'PlayMode', 'PlayVol', 'Medias', 'Terms', 'AreaMasks',
    ##           'Groups', 'PowerAheadPlay']
    nst='2023-10-10 '+starttime
    gr=area_to_num()
    onerec=[name,2,0,0,nst,'2023-10-10',65663,1,0,0,
            medias,'','',gr,0]
    data.append(onerec)
    play_area=[]

def clear_checkbox():
    ui.checkBox.setChecked(False)
    ui.checkBox_2.setChecked(False)
    ui.checkBox_3.setChecked(False)
    ui.checkBox_4.setChecked(False)
    ui.checkBox_5.setChecked(False)
    ui.checkBox_6.setChecked(False)
    ui.checkBox_7.setChecked(False)

## 点击添加按钮触发的函数
def add_btn():
    global name
    global medias
    global lasttime
    global onerec
    global starttime
    global tag
    global play_area
    global ntt
    getallvalues()
    ## 如果没有输入任何的内容
    if play_area==[]   or starttime=='00:00:00' or medias=='' or name==''  :
        QMessageBox.critical(None, "错误", "有信息漏填")
        tag=0
        play_area=[]
    else:
        ### 获取每个项目的内容，显示出来在table
        endtime=changestrtotime(starttime)
        settablevalue(endtime)
        ## 生成一个列表，包含每条记录的内容
        print(play_area)
        generate_data()

        ## 重置所有的内容
        name=''
        medias=''
        starttime=''

        lasttime=0
        onerec=[]
        play_area=[]
        tag=0
        ntt=''
        # 使得原本的内容清空
        ui.lineEdit.setText('')
        ui.lineEdit_2.setText('')
        clear_checkbox()




def save_excel():
    global data
    workbook = xlwt.Workbook(encoding='utf-8')  #rec_gb_ex_data
    sheet = workbook.add_sheet('rec_gb-sheet')
    columns = ['Name', 'JobType', 'JobMask', 'Duration', 'StartTime', 'StopTime',
              'JobData', 'RepeatNum', 'PlayMode', 'PlayVol', 'Medias', 'Terms', 'AreaMasks',
              'Groups', 'PowerAheadPlay']
    for col, column in enumerate(columns):
        sheet.write(0, col, column)
    for i, row in enumerate(data):
        for j, colval in enumerate(row):
            sheet.write(i + 1, j, colval)

    workbook.save('./rec_gb_ex_data.xls')


############################################################
app = QApplication(sys.argv)

win = QMainWindow()
#
# win.setGeometry(400, 400, 400, 300)
#
# win.setWindowTitle("Pyqt5 Tutorial")
#
# win.show()



ui=wm.Ui_MainWindow()
ui.setupUi(win)
win.show()
ui.pushButton.clicked.connect(add_btn)
ui.pushButton_3.clicked.connect(readfile)
ui.pushButton_2.clicked.connect(save_excel)


sys.exit( app.exec_() )
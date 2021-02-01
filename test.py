import wave
from tkinter import filedialog
from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtCore import pyqtSlot, QUrl
from PyQt5.QtMultimedia import QMediaPlaylist, QMediaPlayer, QMediaContent
from PyQt5.QtWidgets import *
import pandas as pd
from tkinter import *
import sys
import pyqtgraph as pg

class Ui_MainWindow(QWidget):
    def __init__(self):
        super().__init__()

    def setupUi(self, MainWindow):
        self.soundPath = '' #음원 경로
        self.absolutePath = False #절대경로
        self.relativePath = False #상대경로
        self.startTime = 0

        #Player
        self.player = QMediaPlayer()
        self.relplaylist = QMediaPlaylist()

        #main
        self.playlist = []
        self.selectedList = [0]
        self.playOption = QMediaPlaylist.Sequential

        # 프로그램 창
        MainWindow.setObjectName("MainWindow")
        MainWindow.resize(1078, 797)
        self.centralwidget = QtWidgets.QWidget(MainWindow)
        self.centralwidget.setObjectName("centralwidget")

        # QtableWidget(엑셀)
        self.excelWidget = QtWidgets.QTableWidget(self.centralwidget)
        self.excelWidget.setGeometry(QtCore.QRect(10, 10, 1061, 481))
        self.excelWidget.setObjectName("excelWidget")

        # 재생/일시정지
        self.playBtn = QtWidgets.QPushButton(self.centralwidget)
        self.playBtn.setGeometry(QtCore.QRect(100, 500, 71, 71))
        self.playBtn.setObjectName("playBtn")

        # 되감기 2초전
        self.rewindBtn = QtWidgets.QPushButton(self.centralwidget)
        self.rewindBtn.setGeometry(QtCore.QRect(10, 500, 71, 71))
        self.rewindBtn.setObjectName("rewindBtn")
        self.rewindBtn.setShortcut("F6")

        # 멈춤
        self.stopBtn = QtWidgets.QPushButton(self.centralwidget)
        self.stopBtn.setGeometry(QtCore.QRect(190, 500, 71, 71))
        self.stopBtn.setObjectName("stopBtn")

        # 빨리감기 2초후
        self.fowardBtn = QtWidgets.QPushButton(self.centralwidget)
        self.fowardBtn.setGeometry(QtCore.QRect(280, 500, 71, 71))
        self.fowardBtn.setObjectName("fowardBtn")

        # 플레이어 시간 슬라이더
        self.horizontalSlider = QtWidgets.QSlider(self.centralwidget)
        self.horizontalSlider.setGeometry(QtCore.QRect(370, 500, 701, 71))
        self.horizontalSlider.setOrientation(QtCore.Qt.Horizontal)
        self.horizontalSlider.setObjectName("horizontalSlider")

        # 플레이어 시간 슬라이더 현재 진행 time
        self.playTime = QtWidgets.QLabel(self.centralwidget)
        self.playTime.setGeometry(QtCore.QRect(365, 580, 61, 41))
        font = QtGui.QFont()
        font.setFamily("Agency FB")
        font.setPointSize(20)
        self.playTime.setFont(font)
        self.playTime.setObjectName("playTime")

        # 플레이어 시간 슬라이더 전체 time
        self.endTime = QtWidgets.QLabel(self.centralwidget)
        self.endTime.setGeometry(QtCore.QRect(1020, 580, 61, 41))
        font = QtGui.QFont()
        font.setFamily("Agency FB")
        font.setPointSize(20)
        self.endTime.setFont(font)
        self.endTime.setObjectName("endTime")

        # rewind 버튼 단축 Label
        self.rewindLabel = QtWidgets.QLabel(self.centralwidget)
        self.rewindLabel.setGeometry(QtCore.QRect(35, 580, 61, 41))
        font = QtGui.QFont()
        font.setFamily("Agency FB")
        font.setPointSize(20)
        self.rewindLabel.setFont(font)
        self.rewindLabel.setObjectName("rewindLabel")

        # play 버튼 단축 Label
        self.playLabel = QtWidgets.QLabel(self.centralwidget)
        self.playLabel.setGeometry(QtCore.QRect(125, 580, 61, 41))
        font = QtGui.QFont()
        font.setFamily("Agency FB")
        font.setPointSize(20)
        self.playLabel.setFont(font)
        self.playLabel.setObjectName("playLabel")

        # stop 버튼 단축 Label
        self.stopLabel = QtWidgets.QLabel(self.centralwidget)
        self.stopLabel.setGeometry(QtCore.QRect(215, 580, 61, 41))
        font = QtGui.QFont()
        font.setFamily("Agency FB")
        font.setPointSize(20)
        self.stopLabel.setFont(font)
        self.stopLabel.setObjectName("stopLabel")

        # foward 버튼 단축 Label
        self.fowardLabel = QtWidgets.QLabel(self.centralwidget)
        self.fowardLabel.setGeometry(QtCore.QRect(305, 580, 61, 41))
        font = QtGui.QFont()
        font.setFamily("Agency FB")
        font.setPointSize(20)
        self.fowardLabel.setFont(font)
        self.fowardLabel.setObjectName("fowardLabel")

        #경로설정
        self.comboBox = QtWidgets.QComboBox(self.centralwidget)
        self.comboBox.setGeometry(QtCore.QRect(10, 680, 111, 31))
        self.comboBox.setObjectName("comboBox")

        # wav 파형
        self.openGLWidget = QtWidgets.QOpenGLWidget(self.centralwidget)
        self.openGLWidget.setGeometry(QtCore.QRect(370, 630, 690, 120))
        self.openGLWidget.setObjectName("openGLWidget")
        MainWindow.setCentralWidget(self.centralwidget)

        # 상대경로 / 절대경로 표시해주는 label
        self.nowPath = QtWidgets.QLabel(self.centralwidget)
        self.nowPath.setGeometry(QtCore.QRect(10, 720, 270, 41))
        font = QtGui.QFont()
        font.setFamily("Agency FB")
        font.setPointSize(20)
        self.nowPath.setFont(font)
        self.nowPath.setObjectName("nowPath")

        MainWindow.setCentralWidget(self.centralwidget)

        # 메뉴바
        self.menubar = QtWidgets.QMenuBar(MainWindow)
        self.menubar.setGeometry(QtCore.QRect(0, 0, 1078, 26))
        self.menubar.setObjectName("menubar")

        # 메뉴
        self.menu = QtWidgets.QMenu(self.menubar)
        self.menu.setObjectName("menu")

        MainWindow.setMenuBar(self.menubar)

        # 상태바
        self.statusbar = QtWidgets.QStatusBar(MainWindow)
        self.statusbar.setObjectName("statusbar")

        MainWindow.setStatusBar(self.statusbar)

        # 메뉴 - 엑셀불러오기
        self.ActionExcelLoad = QtWidgets.QAction(MainWindow)
        self.ActionExcelLoad.setObjectName("ActionExcelLoad")

        # 메뉴 - 상대경로
        self.ActionSoundPath = QtWidgets.QAction(MainWindow)
        self.ActionSoundPath.setObjectName("ActionSoundPath")

        # 메뉴 - 절대경로
        self.AbsoulteSoundPath = QtWidgets.QAction(MainWindow)
        self.AbsoulteSoundPath.setObjectName("AbsoulteSoundPath")

        # 엑셀 저장
        self.AactionSave = QtWidgets.QAction(MainWindow)
        self.AactionSave.setObjectName("AactionSave")

        # 프로그램 종료
        self.AactionExit = QtWidgets.QAction(MainWindow)
        self.AactionExit.setObjectName("AactionExit")

        # 상단 메뉴에 있는 각각의 버튼 리스트
        self.menu.addAction(self.ActionExcelLoad)
        self.menu.addAction(self.ActionSoundPath)
        self.menu.addAction(self.AbsoulteSoundPath)
        self.menu.addSeparator()
        self.menu.addAction(self.AactionSave)
        self.menu.addAction(self.AactionExit)
        self.menubar.addAction(self.menu.menuAction())

        # 기능들 UI
        self.retranslateUi(MainWindow)
        QtCore.QMetaObject.connectSlotsByName(MainWindow)

        self.ActionExcelLoad.triggered.connect(self.ActionExcelLoad_clicked)
        self.ActionSoundPath.triggered.connect(self.ActionSoundPath_clicked)
        self.AbsoulteSoundPath.triggered.connect(self.AbsoulteSoundPath_clicked)
        #self.excelWidget.cellClicked.connect(self.excelWidget_cellClicked)
        self.AactionSave.triggered.connect(self.ActionSave_clicked)
        self.playBtn.clicked.connect(self.playBtn_clicked)
        self.stopBtn.clicked.connect(self.stopBtn_clicked)
        self.rewindBtn.clicked.connect(self.rewindBtn_clicked)
        self.fowardBtn.clicked.connect(self.fowardBtn_clicked)
        self.AactionExit.triggered.connect(self.AactionExit_clicked)
        self.excelWidget.currentCellChanged.connect(self.cellChanged)
        self.player.durationChanged.connect(self.durationChanged)
        self.player.positionChanged.connect(self.positionChanged)

        self.rewindBtn.setShortcut('F5')
        self.playBtn.setShortcut('F6')
        self.stopBtn.setShortcut('F7')
        self.fowardBtn.setShortcut('F8')

    def retranslateUi(self, MainWindow):
        _translate = QtCore.QCoreApplication.translate
        MainWindow.setWindowTitle(_translate("MainWindow", "TRtool"))
        self.playBtn.setText(_translate("MainWindow", "▶"))
        self.rewindBtn.setText(_translate("MainWindow", "◀◀"))
        self.stopBtn.setText(_translate("MainWindow", "■"))
        self.fowardBtn.setText(_translate("MainWindow", "▶▶"))
        self.playTime.setText(_translate("MainWindow", "00:00"))
        self.endTime.setText(_translate("MainWindow", "00:00"))
        self.rewindLabel.setText(_translate("MainWindow", "F5"))
        self.playLabel.setText(_translate("MainWindow", "F6"))
        self.stopLabel.setText(_translate("MainWindow", "F7"))
        self.fowardLabel.setText(_translate("MainWindow", "F8"))
        self.nowPath.setText(_translate("MainWindow", "경로를 설정하세요"))
        self.menu.setTitle(_translate("MainWindow", "메뉴"))
        self.ActionExcelLoad.setText(_translate("MainWindow", "엑셀 파일 가져오기"))
        self.ActionSoundPath.setText(_translate("MainWindow", "음원 상대경로 지정"))
        self.AbsoulteSoundPath.setText(_translate("MainWindow", "음원 절대경로 지정"))
        self.AactionSave.setText(_translate("MainWindow", "저장(S)"))
        self.AactionExit.setText(_translate("MainWindow", "종료(Q)"))

    # cell이 바뀔때마다 셀 클릭
    def cellChanged(self):
        try:
            self.excelWidget_cellClicked()
        except Exception as E:
            print(E)

    # 불러오기시 xlsx 형식의 확장자가 아니면 경고 메세지 출력 후 로드되지 않게 설정(어떤 이유에서 인지 알집을 불러오기시 멈춤현상 발견)
    def ActionExcelLoad_clicked(self):
        try:
            loadFileName = QFileDialog.getOpenFileName(self, 'Open file', "",
                                        "All Files(*);; Excel Files(*.xlsx)", '/home')
            if loadFileName[0] and loadFileName[0][-4:] == 'xlsx':
                self.sample = pd.read_excel(loadFileName[0])
                self.excelWidgetUI(self.sample)
            else:
                QMessageBox.about(self, "Warning", "파일을 선택하지 않았거나 Excel형식이 아닙니다.")
        except Exception as E:
            print(E)

    # 상대경로 음원 설정
    def ActionSoundPath_clicked(self):
        self.soundPath = str(QtWidgets.QFileDialog.getExistingDirectory())
        if self.soundPath:
            self.relativePath = True #상대경로 true
            self.absolutePath = False #절대경로 False
            QMessageBox.about(self, "음원 상대경로 지정", "상대경로 : " + self.soundPath + "\n가 지정되었습니다.")
            self.nowPath.setText(QtCore.QCoreApplication.translate("MainWindow", "상대경로"))
        else:
            QMessageBox.about(self, "Warning", "경로가 설정되지 않았습니다.")

    # 절대경로 음원 설정
    def AbsoulteSoundPath_clicked(self):
        self.relativePath = False  # 상대경로 False
        self.absolutePath = True  # 절대경로 true
        self.nowPath.setText(QtCore.QCoreApplication.translate("MainWindow", "절대경로"))
        QMessageBox.about(self, "음원 절대경로 지정", "절대경로로 설정되었습니다.")

    # 메뉴바의 저장하기 클릭
    def ActionSave_clicked(self):
        # 취소하거나 X를 클릭하게 되면 해당 경고 메세지가 나타남
        if self.excelWidget.item(0,0) == None :
            QMessageBox.about(self, "Warning", "엑셀이 선택되지 않았습니다.")
        else:
            savedic = {}
            savelist = []

            for columns in range(self.excelWidget.columnCount()):
                for row in range(self.excelWidget.rowCount()):
                    savelist.append(self.excelWidget.item(row, columns).text())
                savedic[self.sample.columns.ravel().item(columns)] = savelist
                savelist = []

            df = pd.DataFrame(savedic)
            try:
                self.Save(df)
            except Exception as E:
                print(E)

    # 다른이름으로 저장
    def Save(self, df):
        root = Tk().withdraw()
        savefilename = filedialog.asksaveasfilename(title='Select file',
                                                    filetypes=(('Excel files', '*.xlsx'), ('all files', '*.*')),
                                                    initialfile='noname.xlsx')
        if savefilename[0]:
            df.to_excel(savefilename, sheet_name='TR', index=False)
        else:
            QMessageBox.about(self, "Warning", "취소하였습니다.")

    # 엑셀파일 불러오기시 Qtable에 엑셀형태로 데이터를 설정
    def excelWidgetUI(self, sample):
        self.comboBox.clear()
        for i in range(len(sample.columns.ravel())):
            self.comboBox.addItem("")
            self.comboBox.setItemText(i, QtCore.QCoreApplication.translate("MainWindow", sample.columns.ravel().item(i)))
        self.excelWidget.setRowCount(len(sample))
        self.excelWidget.setColumnCount(len(sample.columns.ravel()))

        for rows in range(len(sample)):
            for columns in range(len(sample.columns.ravel())):
                 self.excelWidget.setItem(rows, columns, QTableWidgetItem(str(sample[sample.columns.ravel().item(columns)][rows])))

    # 셀 클릭시 해당 열의 wav파일 재생
    @pyqtSlot(int, int)
    def excelWidget_cellClicked(self):
        try:
            if self.absolutePath == True:
                wavdir = self.excelWidget.item(self.excelWidget.currentIndex().row(), self.comboBox.currentIndex()).text()
                self.wavPlay(wavdir)
            elif self.relativePath == True:
                wavdir = self.soundPath + "/" + self.excelWidget.item(self.excelWidget.currentIndex().row(), self.comboBox.currentIndex()).text()
                self.wavPlay(wavdir)
            else:
                QMessageBox.about(self, "Warning", "상대경로 또는 절대경로를 설정해주세요.")
        except Exception as E:
            print(E)

    # wav를 실행함
    def wavPlay(self, wavdir):
        self.player.stop()
        try:
            self.playlist.clear()
            self.playlist.append(wavdir)
            self.play(self.playlist, self.selectedList[0], self.playOption)

        except Exception as E:
            if str(E) == "file does not start with RIFF id":
                QMessageBox.about(self, "Warning", str(E) + "\n해당 wav는 RIFF 형식이 아닙니다.")
            else:
                QMessageBox.about(self, "Warning", str(E) + "\n해당 경로에 wav 파일이 존재하지 않습니다")

    # 음원 play
    def play(self, playlists, startRow=0, option=QMediaPlaylist.Sequential):
        if self.player.state() == QMediaPlayer.PausedState:
            self.player.play()
        else:
            self.createPlaylist(playlists, startRow, option)
            self.player.setPlaylist(self.relplaylist)
            self.relplaylist.setCurrentIndex(startRow)
            self.player.play()

    # 음원의 경로를 불러와서 쓰레드처럼 background에서 돌게함
    def createPlaylist(self, playlists, startRow=0, option=QMediaPlaylist.Sequential):
        self.relplaylist.clear()

        for path in playlists:
            url = QUrl.fromLocalFile(path)
            self.relplaylist.addMedia(QMediaContent(url))

        self.relplaylist.setPlaybackMode(option)

    # '▶' 버튼클릭시 음원이 플레이중이면 pause 상태로, 음원이 pause 상태이면 계속 진행
    def playBtn_clicked(self):
        if self.player.state() == QMediaPlayer.PlayingState:
            self.player.pause()
        else:
            self.player.play()

    # '■' 버튼 클릭시 정지
    def stopBtn_clicked(self):
        self.player.stop()
        self.playTime.setText(QtCore.QCoreApplication.translate("MainWindow","00:00"))

    # '◀◀' 버튼 클릭시 2초 전으로 되감기
    def rewindBtn_clicked(self):
        try:
            self.player.setPosition(self.player.position() - 2000)
        except Exception as E:
            print(E)

    # '▶▶' 버튼 클릭시 2초 후로 빨리감기
    def fowardBtn_clicked(self):
        try:
            self.player.setPosition(self.player.position() + 2000)
        except Exception as E:
            print(E)

    # 메뉴바의 종료 버튼
    def AactionExit_clicked(self):
        sys.exit()

    # cell을 다른걸로 클릭하여 음원이 바뀔때 총 길이를 바꿔줌
    def durationChanged(self, msec):
        if msec > -1:
            self.updateDurationChanged(self.relplaylist, msec)

    # cell을 다른걸로 클릭하여 음원이 바뀔때 현재 진행되는 시간을 바꿔줌
    def positionChanged(self, msec):
        if msec > -1:
            self.updatePositionChanged(self.relplaylist, msec)

    # 음원의 총 길이를 설정하여 endtime 표시
    def updateDurationChanged(self, index, msec):
        if msec != 0:
            m, s = divmod(msec/1000, 60)
            self.endTime.setText(QtCore.QCoreApplication.translate("MainWindow",
                                                                   "%02d:%02d" % (m, s)))
        self.horizontalSlider.setRange(0, msec)

    # 음원의 현재 진행시간을 실시간으로 설정하여 playtime 표시
    def updatePositionChanged(self, index, msec):
        if msec != 0:
            m, s = divmod(msec / 1000, 60)
            self.playTime.setText(QtCore.QCoreApplication.translate("MainWindow",
                                                                   "%02d:%02d" % (m, s)))
        self.horizontalSlider.setValue(msec)

if __name__ == "__main__":
    app = QtWidgets.QApplication(sys.argv)
    MainWindow = QtWidgets.QMainWindow()
    ui = Ui_MainWindow()
    ui.setupUi(MainWindow)
    MainWindow.show()
    sys.exit(app.exec_())


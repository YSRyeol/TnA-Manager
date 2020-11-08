from PyQt5.QtWidgets import *
from PyQt5.QtGui import *
from PyQt5.QtCore import *
import datetime
import os
import sys
import time

if not os.path.isdir('Log'):
    os.makedirs(os.path.join('Log'))
if not os.path.isdir('Excel'):
    os.makedirs(os.path.join('Excel'))

curYear = datetime.datetime.today().year
curMonth = datetime.datetime.today().month
curDay = datetime.datetime.today().day
day_of_month = {
        1: 31, 2: 28, 3: 31, 4: 30, 5: 31, 6: 30,
        7: 31, 8: 31, 9: 30, 10: 31, 11: 30, 12: 31
    }

times = sorted(['0' + str(i) + ':00' if i < 10 else str(i) + ':00' for i in range(7, 24)] + ['0' + str(i) + ':30' if i < 10 else str(i) + ':30' for i in range(7, 24)])[:-1]

class AlertDialog(QDialog):
    def __init__(self, m, dataDict=None):
        super().__init__()
        self.setWindowTitle('알림')
        self.setContentsMargins(10, 10, 10, 10)

        self.MENU = m
        self.DATA_DICT = dataDict
        self.BEFORE_WIDTH = 0
        self.BEFORE_HEIGHT = 0

        self.setupUI()

    def setupUI(self):
        self.layoutDialog = QVBoxLayout()

        if self.MENU == 'DATA_IS_NOT_VALID':
            self.setMsgLayout('데이터를 올바르게 입력되지 않았습니다.\n확인 후 다시 시도해 주세요.')
            self.setButtonLayout()
        elif self.MENU == 'FILE_IS_NOT_FOUND':
            self.setMsgLayout('해당 엑셀 파일이 존재하지 않습니다.\n생성할까요?')
            self.setButtonLayout(True)
        elif self.MENU == 'FILE_IS_OPENED':
            self.setMsgLayout('해당 엑셀 파일이 열려있습니다.\n엑셀 파일 종료 후 다시 시도해 주세요.')
            self.setButtonLayout()
        elif self.MENU == 'SHEET_IS_NOT_FOUND':
            self.setMsgLayout('해당 시트가 존재하지 않습니다.\n생성할까요?')
            self.setButtonLayout(True)
        elif self.MENU == 'DATA_IS_EXIST':
            self.setMsgLayout('해당 날짜에 이미 아래 데이터가 존재합니다.\n데이터를 덮어쓸까요?')
            self.setDataLayout()
            self.setButtonLayout(True)
        elif self.MENU == 'ASK_ENTER_DATA':
            self.setMsgLayout('아래 데이터를 입력할까요?')
            self.setDataLayout()
            self.setButtonLayout(True)

        self.setLayout(self.layoutDialog)

    def setMsgLayout(self, msg):
        self.layoutDialog.addWidget(QLabel(msg))

    def setDataLayout(self):
        grid = QGridLayout()
        grid.addWidget(QLabel(f"{self.DATA_DICT['date']['y']}-{self.DATA_DICT['date']['m']}-{self.DATA_DICT['date']['d']}"), 0, 0, 1, -1, alignment=Qt.AlignCenter)
        grid.addWidget(QLabel('이름'), 1, 0, alignment=Qt.AlignCenter)
        grid.addWidget(self.setVLine(), 1, 1, -1, 1)
        grid.addWidget(QLabel('근무지'), 1, 2, alignment=Qt.AlignCenter)
        grid.addWidget(self.setVLine(), 1, 3, -1, 1)
        grid.addWidget(QLabel('출퇴근 시각'), 1, 4, alignment=Qt.AlignCenter)
        grid.addWidget(self.setVLine(), 1, 5, -1, 1)
        grid.addWidget(QLabel('근무 시각'), 1, 6, alignment=Qt.AlignCenter)
        grid.addWidget(self.setVLine(), 1, 7, -1, 1)
        grid.addWidget(QLabel('근무 시간'), 1, 8, alignment=Qt.AlignCenter)
        grid.addWidget(self.setVLine(), 1, 9, -1, 1)
        grid.addWidget(QLabel('잔업 시간'), 1, 10, alignment=Qt.AlignCenter)
        grid.addWidget(self.setVLine(), 1, 11, -1, 1)
        grid.addWidget(QLabel('특근 시간'), 1, 12, alignment=Qt.AlignCenter)
        grid.addWidget(self.setVLine(), 1, 13, -1, 1)
        grid.addWidget(self.setHLine(), 2, 0, 1, -1)
        grid.addWidget(QLabel(self.DATA_DICT['name']), 3, 0, -1, 1, alignment=Qt.AlignCenter)
        grid.addWidget(QLabel('본사'), 3, 2, alignment=Qt.AlignCenter)
        grid.addWidget(QLabel(self.DATA_DICT['totalWorking']), 3, 4, -1, 1, alignment=Qt.AlignCenter)
        if self.DATA_DICT['hq'] is not None:
            grid.addWidget(QLabel(str(self.DATA_DICT['hq'][0])), 3, 6, alignment=Qt.AlignCenter)
            grid.addWidget(QLabel(str(self.DATA_DICT['hq'][1])), 3, 8, alignment=Qt.AlignCenter)
            grid.addWidget(QLabel(str(self.DATA_DICT['hq'][2]) if self.DATA_DICT['hq'][2] is not None else ''), 3, 10, alignment=Qt.AlignCenter)
            grid.addWidget(QLabel(str(self.DATA_DICT['hq'][3]) if self.DATA_DICT['hq'][3] is not None else ''), 3, 12, alignment=Qt.AlignCenter)
        grid.addWidget(self.setHLine(), 4, 2)
        grid.addWidget(self.setHLine(), 4, 6, 1, -1)
        grid.addWidget(QLabel('400'), 5, 2, alignment=Qt.AlignCenter)
        if self.DATA_DICT['400'] is not None:
            grid.addWidget(QLabel(str(self.DATA_DICT['400'][0])), 5, 6, alignment=Qt.AlignCenter)
            grid.addWidget(QLabel(str(self.DATA_DICT['400'][1])), 5, 8, alignment=Qt.AlignCenter)
            grid.addWidget(QLabel(str(self.DATA_DICT['400'][2]) if self.DATA_DICT['400'][2] is not None else ''), 5, 10, alignment=Qt.AlignCenter)
            grid.addWidget(QLabel(str(self.DATA_DICT['400'][3]) if self.DATA_DICT['400'][3] is not None else ''), 5, 12, alignment=Qt.AlignCenter)
        grid.addWidget(self.setHLine(), 6, 2)
        grid.addWidget(self.setHLine(), 6, 6, 1, -1)
        grid.addWidget(QLabel('어비리'), 7, 2, alignment=Qt.AlignCenter)
        if self.DATA_DICT['eobiri'] is not None:
            grid.addWidget(QLabel(str(self.DATA_DICT['eobiri'][0])), 7, 6, alignment=Qt.AlignCenter)
            grid.addWidget(QLabel(str(self.DATA_DICT['eobiri'][1])), 7, 8, alignment=Qt.AlignCenter)
            grid.addWidget(QLabel(str(self.DATA_DICT['eobiri'][2]) if self.DATA_DICT['eobiri'][2] is not None else ''), 7, 10, alignment=Qt.AlignCenter)
            grid.addWidget(QLabel(str(self.DATA_DICT['eobiri'][3]) if self.DATA_DICT['eobiri'][3] is not None else ''), 7, 12, alignment=Qt.AlignCenter)

        self.layoutDialog.addLayout(grid)

    def setButtonLayout(self, isChoice=False):
        btnLayout = QHBoxLayout()
        btnLayout.setSpacing(5)
        btnLayout.setContentsMargins(0, 10, 0, 0)
        btnLayout.addStretch()
        if isChoice:
            btnConfirm = QPushButton('예')
            btnConfirm.clicked.connect(self.accept)
            btnCancel = QPushButton('아니오')
            btnCancel.clicked.connect(self.reject)

            btnLayout.addWidget(btnConfirm)
            btnLayout.addWidget(btnCancel)
        else:
            btnConfirm = QPushButton('확인')
            btnConfirm.clicked.connect(self.reject)

            btnLayout.addWidget(btnConfirm)
        btnLayout.addStretch()

        self.layoutDialog.addLayout(btnLayout)

    def setVLine(self):
        vLine = QFrame()
        vLine.setFrameShape(QFrame.VLine)
        vLine.setFrameShadow(QFrame.Sunken)

        return vLine

    def setHLine(self):
        hLine = QFrame()
        hLine.setFrameShape(QFrame.HLine)
        hLine.setFrameShadow(QFrame.Sunken)

        return hLine

class MainWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle('T&A Manager')

        with open('employee.txt', 'r', encoding='utf-8-sig') as f:
            self.EMP_LIST = [emp.replace('\n', '') for emp in f.readlines()]
        self.LOG_LIST = []

        self.setupUI()

    def setSize(self):
        self.setLogLayout()
        self.setFixedSize(self.width(), self.height())

    def setupUI(self):
        self.layoutDialog = QGridLayout()
        self.layoutDialog.setContentsMargins(20, 20, 20, 20)

        self.setDateLayout()
        self.setNameLayout()
        self.setGoToWorkTimeLayout()
        self.setGoHomeTimeLayout()
        self.setWorkingTimeLayout()
        self.setIsDinnerLayout()
        self.setButtonLayout()

        self.setLayout(self.layoutDialog)

    # 날짜 입력
    def setDateLayout(self):
        lbDate = QLabel('날짜:')
        ## 년도
        self.cbYear = QComboBox()
        self.cbYear.addItems([str(i) for i in range(curYear, 1999, -1)])
        self.cbYear.currentTextChanged.connect(self.changedYear)
        ## 월
        self.cbMonth = QComboBox()
        self.cbMonth.addItems([str(i) for i in range(1, 13)])
        self.cbMonth.setCurrentText(str(curMonth))
        self.cbMonth.currentTextChanged.connect(self.changeMonth)
        ## 일
        self.cbDay = QComboBox()
        self.cbDay.addItems(self.getDays())
        self.cbDay.setCurrentText(str(curDay))

        layoutDate = QHBoxLayout()
        layoutDate.addWidget(self.cbYear)
        layoutDate.addWidget(self.cbMonth)
        layoutDate.addWidget(self.cbDay)

        self.layoutDialog.addWidget(lbDate, 0, 0, alignment=Qt.AlignRight)
        self.layoutDialog.addLayout(layoutDate, 0, 1, 1, -1)

    # 이름 입력
    def setNameLayout(self):
        lbName = QLabel('이름:')
        self.cbName = QComboBox()
        self.cbName.addItems(self.EMP_LIST)

        self.layoutDialog.addWidget(lbName, 1, 0, alignment=Qt.AlignRight)
        self.layoutDialog.addWidget(self.cbName, 1, 1)

    # 출근 시간 입력
    def setGoToWorkTimeLayout(self):
        goToWorkTimes = times[:times.index('08:00') + 1]

        lbGoToWorkTime = QLabel('출근 시각:')
        self.cbGoToWorkTime = QComboBox()
        self.cbGoToWorkTime.addItems(goToWorkTimes)
        self.cbGoToWorkTime.setCurrentText('08:00')
        self.cbGoToWorkTime.currentIndexChanged.connect(self.changedTotalWorkingTime)

        self.layoutDialog.addWidget(lbGoToWorkTime, 2, 0, alignment=Qt.AlignRight)
        self.layoutDialog.addWidget(self.cbGoToWorkTime, 2, 1)

    # 퇴근 시간 입력
    def setGoHomeTimeLayout(self):
        goHomeTimes = times[times.index('17:00'):]

        lbGoHomeTime = QLabel('퇴근 시각:')
        self.cbGoHomeTime = QComboBox()
        self.cbGoHomeTime.addItems(goHomeTimes)
        self.cbGoHomeTime.currentIndexChanged.connect(self.changedTotalWorkingTime)

        self.layoutDialog.addWidget(lbGoHomeTime, 3, 0, alignment=Qt.AlignRight)
        self.layoutDialog.addWidget(self.cbGoHomeTime, 3, 1)

    # 근무지별 근무 시간 입력
    def setWorkingTimeLayout(self):
        # 라벨
        lbPlace = QLabel()
        lbStartTime = QLabel('시작 시각:')
        lbEndTime = QLabel('종료 시각:')
        
        # 본사
        lbHq = QLabel('본사')
        lbHq.setAlignment(Qt.AlignHCenter)
        self.cbStartTimeHq = QComboBox()
        self.cbStartTimeHq.addItems(self.getWorkingTimes())
        self.cbEndTimeHq = QComboBox()
        self.cbEndTimeHq.addItems(self.getWorkingTimes())
        
        # 400번지
        lb400 = QLabel('400번지')
        lb400.setAlignment(Qt.AlignHCenter)
        self.cbStartTime400 = QComboBox()
        self.cbStartTime400.addItems(self.getWorkingTimes())
        self.cbEndTime400 = QComboBox()
        self.cbEndTime400.addItems(self.getWorkingTimes())
        
        # 어비리
        lbEobiri = QLabel('어비리')
        lbEobiri.setAlignment(Qt.AlignHCenter)
        self.cbStartTimeEobiri = QComboBox()
        self.cbStartTimeEobiri.addItems(self.getWorkingTimes())
        self.cbEndTimeEobiri = QComboBox()
        self.cbEndTimeEobiri.addItems(self.getWorkingTimes())
        
        self.layoutDialog.addWidget(lbPlace, 4, 0)
        self.layoutDialog.addWidget(lbStartTime, 5, 0)
        self.layoutDialog.addWidget(lbEndTime, 6, 0)
        self.layoutDialog.addWidget(lbHq, 4, 1)
        self.layoutDialog.addWidget(self.cbStartTimeHq, 5, 1)
        self.layoutDialog.addWidget(self.cbEndTimeHq, 6, 1)
        self.layoutDialog.addWidget(lb400, 4, 2)
        self.layoutDialog.addWidget(self.cbStartTime400, 5, 2)
        self.layoutDialog.addWidget(self.cbEndTime400, 6, 2)
        self.layoutDialog.addWidget(lbEobiri, 4, 3)
        self.layoutDialog.addWidget(self.cbStartTimeEobiri, 5, 3)
        self.layoutDialog.addWidget(self.cbEndTimeEobiri, 6, 3)

    # 저녁식사 확인
    def setIsDinnerLayout(self):
        self.cbIsDinner = QCheckBox("저녁식사")
        
        self.layoutDialog.addWidget(self.cbIsDinner, 3, 2, 1, -1, alignment=Qt.AlignRight)

    # 버튼
    def setButtonLayout(self):
        self.pbConfirm = QPushButton('입력')
        self.pbConfirm.clicked.connect(self.inputData)
        self.pbClose = QPushButton('닫기')
        self.pbClose.clicked.connect(QApplication.instance().quit)

        self.layoutDialog.addWidget(QLabel(), 7, 0)
        self.layoutDialog.addWidget(self.pbConfirm, 8, 0, 1, 2, alignment=Qt.AlignRight)
        self.layoutDialog.addWidget(self.pbClose, 8, 2, 1, 2, alignment=Qt.AlignLeft)
        self.layoutDialog.addWidget(QLabel(), 9, 0)

    # 로그
    def setLogLayout(self):
        self.logView = QScrollArea()
        self.logView.setFixedSize(self.width() - 40, 200)
        self.logView.setWidgetResizable(True)
        
        self.logWidget = QWidget()

        self.logBox = QVBoxLayout()
        self.logBox.setAlignment(Qt.AlignTop)
        self.logBox.addWidget(QLabel())

        self.logWidget.setLayout(self.logBox)
        self.logView.setWidget(self.logWidget)

        self.pbRemoveLog = QPushButton('지우기')
        self.pbRemoveLog.clicked.connect(self.logRemove)

        self.layoutDialog.addWidget(self.logView, 10, 0, 1, -1, alignment=Qt.AlignCenter)
        self.layoutDialog.addWidget(self.pbRemoveLog, 11, 0, 1, -1, alignment=Qt.AlignRight)

    # 연도 변경 이벤트
    def changedYear(self):
        selectedDay = self.cbDay.currentText()

        if self.cbMonth.currentIndex() == 1:
            self.cbDay.clear()
            self.cbDay.addItems(self.getDays())

    # 월 변경 이벤트
    def changeMonth(self):
        selectedDay = self.cbDay.currentText()

        self.cbDay.clear()
        self.cbDay.addItems(self.getDays())
        self.cbDay.setCurrentText(selectedDay)

    # 출퇴근 시간 변경 이벤트
    def changedTotalWorkingTime(self):
        self.cbStartTimeHq.clear()
        self.cbEndTimeHq.clear()
        self.cbStartTime400.clear()
        self.cbEndTime400.clear()
        self.cbStartTimeEobiri.clear()
        self.cbEndTimeEobiri.clear()

        self.cbStartTimeHq.addItems(self.getWorkingTimes())
        self.cbEndTimeHq.addItems(self.getWorkingTimes())
        self.cbStartTime400.addItems(self.getWorkingTimes())
        self.cbEndTime400.addItems(self.getWorkingTimes())
        self.cbStartTimeEobiri.addItems(self.getWorkingTimes())
        self.cbEndTimeEobiri.addItems(self.getWorkingTimes())

    # 데이터 입력
    def inputData(self):
        ## 로그
        self.logWrite('데이터 입력 시작')

        curSTHq, curETHq = self.cbStartTimeHq.currentIndex(), self.cbEndTimeHq.currentIndex()
        curST400, curET400 = self.cbStartTime400.currentIndex(), self.cbEndTime400.currentIndex()
        curSTEobiri, curETEobiri = self.cbStartTimeEobiri.currentIndex(), self.cbEndTimeEobiri.currentIndex()

        ## 데이터 유효성 확인
        self.logWrite('데이터 유효성 확인...')
        if bool(curSTHq) != bool(curETHq) or bool(curST400) != bool(curET400) or bool(curSTEobiri) != bool(curETEobiri) or not (curSTHq or curST400 or curSTEobiri):
            self.logWrite('e: 데이터가 유효하지 않습니다.')
            dialog = AlertDialog('DATA_IS_NOT_VALID')
            dialog.exec_()
            del dialog
            self.logWrite('데이터 입력 종료')
            return
        self.logWrite('유효성 확인 완료')

        self.logWrite('데이터 입력 모듈 준비중...')
        from excelManager import ExcelManager
        self.logWrite('모듈 준비 완료')

        year = self.cbYear.currentText()
        ## 엑셀 파일 확인
        self.logWrite(f'{year}.xlsx 파일 확인중...')
        if f'{year}.xlsx' not in os.listdir('Excel/'):
            self.logWrite('i: 해당 파일이 존재하지 않습니다.')
            dialog = AlertDialog('FILE_IS_NOT_FOUND')
            result = dialog.exec_()
            del dialog
            if not result:
                self.logWrite('데이터 입력 종료')
                return
            self.logWrite(f'{year}.xlsx 파일 생성중...')
            em = ExcelManager(year, self.EMP_LIST)
            resultCreate = em.createExcel()
            if resultCreate is not None:
                del em
                self.logWrite(resultCreate, True)
                self.logWrite('데이터 입력 종료')
                return
            self.logWrite('생성 완료')
        else:
            if f'~${year}.xlsx' in os.listdir('Excel/'):
                self.logWrite('e: 해당 파일이 열려 있습니다.')
                dialog = AlertDialog('FILE_IS_OPENED')
                dialog.exec_()
                del dialog
                self.logWrite('데이터 입력 종료')
                return
            self.logWrite(f'{year}.xlsx 파일 여는중...')
            em = ExcelManager(year, self.EMP_LIST)
            resultOpen = em.openExcel()
            if resultOpen is not None:
                del em
                self.logWrite(resultOpen, True)
                self.logWrite('데이터 입력 종료')
                return
            self.logWrite('완료')
        self.logWrite('파일 확인 완료')
        
        month = self.cbMonth.currentText()
        ## 시트 파일 확인
        self.logWrite(f'{month} 시트 확인중...')
        sheets = em.getSheets()
        if not sheets[0]:
            del em
            self.logWrite(sheets[1], True)
            self.logWrite('데이터 입력 종료')
            return
        sheets = sheets[1]
        if month not in sheets:
            self.logWrite('i: 해당 시트가 존재하지 않습니다.')
            dialog = AlertDialog('SHEET_IS_NOT_FOUND')
            result = dialog.exec_()
            del dialog
            if not result:
                del em
                self.logWrite('데이터 입력 종료')
                return
            self.logWrite(f'{month} 시트 생성중..')
            resultCreate = em.createSheet(month, self.getDays()[-1])
            if resultCreate is not None:
                del em
                self.logWrite(resultCreate, True)
                self.logWrite('데이터 입력 종료')
                return
            self.logWrite('완료')
        else:
            self.logWrite(f'{month} 시트 여는 중...')
            resultOpen = em.openSheet(month)
            if resultOpen is not None:
                del em
                self.logWrite(resultOpen, True)
                self.logWrite('데이터 입력 종료')
                return
            self.logWrite('완료')
        self.logWrite('시트 확인 완료')

        name = self.cbName.currentText()
        day = self.cbDay.currentText()
        ## 입력된 데이터가 있는지 확인
        self.logWrite(f'시트 내 {name}: {day}일 데이터 존재 여부 확인중...')
        resultGet = em.getData(name, day)
        if not resultGet[0]:
            del em
            self.logWrite(resultGet[1], True)
            self.logWrite('데이터 입력 종료')
            return
        resultGet = resultGet[1]
        if resultGet['totalWorking'] is not None:
            self.logWrite('i: 입력된 데이터가 존재합니다.')
            dialog = AlertDialog('DATA_IS_EXIST', resultGet)
            result = dialog.exec_()
            del dialog
            if not result:
                del em
                self.logWrite('데이터 입력 종료')
                return
        self.logWrite('데이터 확인 완료')

        ## dataDict 생성
        self.logWrite('입력할 데이터 사전 생성중...')
        totalWorkingTime = self.getWorkingTimes()
        launchStart, launchEnd = totalWorkingTime.index('12:00'), totalWorkingTime.index('13:00')
        launchPlace = 'hq' if curSTHq <= launchStart and curETHq >= launchEnd else '400' if curST400 <= launchStart and curET400 >= launchStart else 'eobiri' if curSTEobiri <= launchStart and curETEobiri >= launchStart else None
        isDinner = True if self.cbIsDinner.isChecked() and curSTHq <= totalWorkingTime.index('17:30') and curETHq >= totalWorkingTime.index('16:00') else False
        isHolyday = True if 5 <= datetime.date(int(year), int(month), int(day)).weekday() else False
        overTimeIndex = totalWorkingTime.index('17:00')
        dataDict = {
            'name': self.cbName.currentText(),
            'date': {'y': year, 'm': month, 'd': day},
            'totalWorking': f'{self.cbGoToWorkTime.currentText()} - {self.cbGoHomeTime.currentText()}',
            'hq': [
                f'{self.cbStartTimeHq.currentText()} - {self.cbEndTimeHq.currentText()}',
                (curETHq - curSTHq - (2 if launchPlace == 'hq' else 0) - (1 if isDinner else 0)) * 0.5,
                (curETHq - (overTimeIndex if overTimeIndex > curSTHq else curSTHq) - (1 if isDinner else 0)) * 0.5 if not isHolyday and curETHq > overTimeIndex else None,
                (curETHq - curSTHq - (2 if launchPlace == 'hq' else 0) - (1 if isDinner else 0)) * 0.5 if isHolyday else None
            ] if curSTHq else None,
            '400': [
                f'{self.cbStartTime400.currentText()} - {self.cbEndTime400.currentText()}',
                (curET400 - curST400 - (2 if launchPlace == '400' else 0)) * 0.5,
                (curET400 - (overTimeIndex if overTimeIndex > curST400 else curST400)) * 0.5 if not isHolyday and curET400 > overTimeIndex else None,
                (curET400 - curST400 - (2 if launchPlace == '400' else 0)) * 0.5 if isHolyday else None
            ] if curST400 else None,
            'eobiri': [
                f'{self.cbStartTimeEobiri.currentText()} - {self.cbEndTimeEobiri.currentText()}',
                (curETEobiri - curSTEobiri - (2 if launchPlace == 'eobiri' else 0)) * 0.5,
                (curETEobiri - (overTimeIndex if overTimeIndex > curSTEobiri else curSTEobiri)) * 0.5 if not isHolyday and curETEobiri > overTimeIndex else None,
                (curETEobiri - curSTEobiri - (2 if launchPlace == 'eobiri' else 0)) * 0.5 if isHolyday else None
            ] if curSTEobiri else None
        }
        self.logWrite('데이터 사전 생성 완료')

        ## 데이터 입력 여부 확인
        dialog = AlertDialog('ASK_ENTER_DATA', dataDict)
        result = dialog.exec_()
        del dialog
        if not result:
            del em
            self.logWrite('데이터 입력 종료')
            return

        ## 엑셀에 데이터 입력
        self.logWrite('데이터 입력중...')
        resultInput = em.inputData(dataDict)
        if resultInput is not None:
            del em
            self.logWrite(resultInput, True)
            self.logWrite('데이터 입력 종료')
            return
        self.logWrite('데이터 입력 완료')
        
        self.logWrite(f'{year}.xlsx 파일 저장중...')
        resultSave = em.saveExcel()
        if resultSave is not None:
            del em
            self.logWrite(resultSave, True)
            self.logWrite('데이터 입력 종료')
            return
        self.logWrite('파일 저장 완료')

        self.resetWidgets()

        self.logWrite()
        self.logWrite('데이터 입력 작업 완료')
        self.logWrite()

    # 위젯 초기화(출퇴근 시간, 근무 시간, 저녁식사 유무)
    def resetWidgets(self):
        self.cbGoToWorkTime.setCurrentText('08:00')
        self.cbGoHomeTime.setCurrentText('17:00')
        self.cbStartTimeHq.setCurrentIndex(0)
        self.cbEndTimeHq.setCurrentIndex(0)
        self.cbStartTime400.setCurrentIndex(0)
        self.cbEndTime400.setCurrentIndex(0)
        self.cbStartTimeEobiri.setCurrentIndex(0)
        self.cbEndTimeEobiri.setCurrentIndex(0)
        self.cbIsDinner.setChecked(False)

    # 로그 출력
    def logWrite(self, log='', fileWrite=False):
        if fileWrite:
            now = time.localtime(time.time())
            fname = f'Log/{curYear}-{curMonth}-{curDay}.log.txt'
            with open(fname, 'a' if os.path.isfile(fname) else 'w') as f:
                f.write(f'[{now.tm_hour}:{now.tm_min}:{now.tm_sec}]\n{log}\n\n')
        lbLog = QLabel(log)
        lbLog.setWordWrap(True)
        self.LOG_LIST.append(lbLog)
        self.logWidget.layout().insertWidget(self.logWidget.layout().count() - 1, lbLog)
        self.logView.verticalScrollBar().setValue(self.logBox.sizeHint().height())

    # 로그 삭제
    def logRemove(self):
        if len(self.LOG_LIST):
            for i in range(len(self.LOG_LIST)):
                log = self.LOG_LIST.pop(-1)
                log.deleteLater()

    # 날짜 리스트
    def getDays(self):
        year = int(self.cbYear.currentText())
        month = int(self.cbMonth.currentText())
        return [str(i) for i in range(1, day_of_month[month] + 2 if month == 2 and ((not year % 4 and year % 100) or not year % 400) else day_of_month[month] + 1)]

    # 근무 시간 리스트
    def getWorkingTimes(self):
        return ['-----'] + times[times.index(self.cbGoToWorkTime.currentText()):times.index(self.cbGoHomeTime.currentText()) + 1]

if __name__ == "__main__":
    try:
        app = QApplication(sys.argv)
        myWindow = MainWindow()
        myWindow.show()
        myWindow.setSize()
        app.exec_()
    except Exception as e:
        now = time.localtime(time.time())
        fname = f'Log/{curYear}-{curMonth}-{curDay}.log.txt'
        with open(fname, 'a' if os.path.isfile(fname) else 'w') as f:
            f.write(f'[{now.tm_hour}:{now.tm_min}:{now.tm_sec}]\n{e}\n\n')
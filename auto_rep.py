from requests_html import HTMLSession
from selenium import webdriver
from selenium.webdriver.firefox.options import Options
from selenium.webdriver.firefox.firefox_binary import FirefoxBinary
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from PyQt5 import QtCore, QtGui, QtWidgets
from multiprocessing import Process
import win32api
import logging
import time
import os
import sys
import json


# URL
URL_DASHBOARD = 'http://ticket.fpt.net/Dashboard/Dashboard/'
URL_LOGIN = 'http://ticket.fpt.net/account/logon'


def main(username, password, browser, message, eventlog):
    init = True
    filename = 'tickets.json'

    while True:
        # get last ticket from ticket.json
        last_ticket = utils.read(filename)['tickets'] if not init else []
        try:
            # Request Dashboard get ticket
            current_ticket = request_ticket(username, password)
            new_ticket = {key: current_ticket[key]
                          for key in set(current_ticket) - set(last_ticket)}
            # notify using win32api
            if len(new_ticket) > 0:
                msg = ', '.join(set(new_ticket))
                eventlog.emit(f'{utils.onTime()} - INFO - {msg}')
                if not init:
                    Process(target=utils.notify, args=(msg,)).start()
            else:
                eventlog.emit(f'{utils.onTime()} - INFO - Nothing')
            # Reply new ticket
            if init is False and current_ticket and len(new_ticket) > 0:
                try:
                    ticket_success = list(reply_ticket(
                        username, password, message, browser, new_ticket))
                    eventlog.emit(f'{utils.onTime()} - INFO - Reply - {", ".join(ticket_success)}')
                    last_ticket.extend(ticket_success)
                except Exception as err:
                    logger.error(f'Reply ticket - {err}')
                    eventlog.emit(f'{utils.onTime()} - ERROR - Reply ticket {err}')
            else:
                if current_ticket:
                    last_ticket.extend(current_ticket)
        except:
            logger.error('Request ticket error')
            eventlog.emit(f'{utils.onTime()} - ERROR - Login error')
        if init:
            init = False
        # save last ticket in tickets.json
        utils.write(filename, {'tickets': last_ticket})
        time.sleep(120)


def request_ticket(username, password):
    """Get url Ticket."""
    data = {'Mail': username, 'Password': password, 'OTP': ''}
    ticket_info = dict()

    session = HTMLSession()
    # if fail retry max 5 time
    for _ in range(5):
        try:
            # Login
            print(data)
            session.post(URL_LOGIN, timeout=20, data=data)
            # Request ticket
            response = session.get(URL_DASHBOARD, timeout=20)
            # Get ticket id and url
            tickets = response.html.find('#SupportNew', first=True)
            tickets = tickets.find('a[target="_blank"]')
            for ticket in tickets:
                ticket_info[ticket.text] = list(ticket.absolute_links)[0]
            return ticket_info
        except Exception as err:
            time.sleep(2)
            logger.error(f'request_ticket {err}')


def reply_ticket(username, password, message, browser, tickets):
    """Send ticket using Selenium."""

    options = Options()
    options.add_argument('--headless')
    try:
        # Setting selenium
        binary = FirefoxBinary(browser)
        driver = webdriver.Firefox(firefox_binary=binary,
                                   executable_path='geckodriver.exe',
                                   options=options)
    except Exception as err:
        logger.error(f'reply_ticket - setup selenium - {err}')
        return

    # Login
    driver.get(URL_LOGIN)
    try:
        driver.find_element_by_id("User").send_keys(username)
        driver.find_element_by_id("Pass").send_keys(password)
        driver.find_element_by_id("btnLogIn").click()
    except Exception as err:
        logger.error(f'reply_ticket - login selenium - {err}')
        return

    # Reply ticket
    for ticket_id in tickets:
        for _ in range(5):
            try:
                driver.get(tickets[ticket_id])
                time.sleep(2)
                WebDriverWait(driver, 30).until(
                    EC.presence_of_element_located(
                        (By.XPATH, "//*[contains(text(), 'chi tiết...')]"))
                )
                elementToClick = driver.find_element_by_xpath(
                    "//*[contains(text(), 'chi tiết...')]")
                location = elementToClick.location
                driver.execute_script("window.scrollTo(0," + str(location["y"]) + ")")
                elementToClick.click()
                time.sleep(2)

                WebDriverWait(driver, 30).until(
                    EC.presence_of_element_located(
                        (By.XPATH, "//div[@aria-labelledby='ui-id-4']/div[11]/div/button[2]"))
                ).click()
                time.sleep(2)
                WebDriverWait(driver, 30).until(
                    EC.presence_of_element_located((By.TAG_NAME, "iframe"))
                )
                driver.switch_to.frame(driver.find_elements_by_tag_name("iframe")[0])
                content = driver.find_element_by_xpath("/html/body/p[1]")
                ActionChains(driver).move_to_element(content).click().send_keys(message).perform()
                driver.switch_to.default_content()
                time.sleep(2)
                WebDriverWait(driver, 30).until(
                    EC.presence_of_element_located((By.ID, "btSend"))
                ).click()
                time.sleep(2)
                yield ticket_id
                break
            except Exception as e:
                logger.error(f'Reply_ticket - {err}')
                time.sleep(3)
    driver.quit()


class utils:
    @staticmethod
    def onTime():
        return time.strftime("%d-%m %H:%M:%S", time.localtime())

    @staticmethod
    def notify(string):
        win32api.MessageBox(0, string, "Alert", 1)

    @staticmethod
    def write(filename, datajson):
        with open(filename, 'w') as f:
            json.dump(datajson, f)

    @staticmethod
    def read(filename):
        try:
            with open(filename, 'r') as f:
                return json.load(f)
        except:
            return None


class Ui_MainWindow(object):
    def setupUi(self, MainWindow):
        MainWindow.setObjectName("MainWindow")
        MainWindow.resize(500, 400)
        self.centralwidget = QtWidgets.QWidget(MainWindow)
        self.centralwidget.setObjectName("centralwidget")
        self.verticalLayout_3 = QtWidgets.QVBoxLayout(self.centralwidget)
        self.verticalLayout_3.setObjectName("verticalLayout_3")
        self.tabWidget = QtWidgets.QTabWidget(self.centralwidget)
        self.tabWidget.setObjectName("tabWidget")
        self.Run = QtWidgets.QWidget()
        self.Run.setObjectName("Run")
        self.verticalLayout = QtWidgets.QVBoxLayout(self.Run)
        self.verticalLayout.setObjectName("verticalLayout")
        self.label = QtWidgets.QLabel(self.Run)
        self.label.setObjectName("label")
        self.verticalLayout.addWidget(self.label)
        self.lineEditUser = QtWidgets.QLineEdit(self.Run)
        self.lineEditUser.setObjectName("lineEditUser")
        self.verticalLayout.addWidget(self.lineEditUser)
        self.lineEditPassword = QtWidgets.QLineEdit(self.Run)
        self.lineEditPassword.setObjectName("lineEditPassword")
        self.verticalLayout.addWidget(self.lineEditPassword)
        self.listWidget = QtWidgets.QListWidget(self.Run)
        self.listWidget.viewport().setProperty("cursor", QtGui.QCursor(QtCore.Qt.IBeamCursor))
        self.listWidget.setObjectName("listWidget")
        item = QtWidgets.QListWidgetItem()
        self.listWidget.addItem(item)
        self.verticalLayout.addWidget(self.listWidget)
        self.buttonBox = QtWidgets.QDialogButtonBox(self.Run)
        self.buttonBox.setStandardButtons(
            QtWidgets.QDialogButtonBox.Cancel | QtWidgets.QDialogButtonBox.Ok)
        self.buttonBox.setObjectName("buttonBox")
        self.verticalLayout.addWidget(self.buttonBox)
        self.tabWidget.addTab(self.Run, "")
        self.setting = QtWidgets.QWidget()
        self.setting.setObjectName("setting")
        self.verticalLayout_2 = QtWidgets.QVBoxLayout(self.setting)
        self.verticalLayout_2.setObjectName("verticalLayout_2")
        self.browserButton = QtWidgets.QPushButton(self.Run)
        self.browserButton.setCursor(QtGui.QCursor(QtCore.Qt.PointingHandCursor))
        self.browserButton.setInputMethodHints(QtCore.Qt.ImhNone)
        self.browserButton.setObjectName("browserButton")
        self.verticalLayout_2.addWidget(self.browserButton)
        self.groupBox = QtWidgets.QGroupBox(self.setting)
        self.groupBox.setTitle("")
        self.groupBox.setObjectName("groupBox")
        self.gridLayout = QtWidgets.QGridLayout(self.groupBox)
        self.gridLayout.setObjectName("gridLayout")
        self.editStartTime = QtWidgets.QDateTimeEdit(self.groupBox)
        self.editStartTime.setCursor(QtGui.QCursor(QtCore.Qt.PointingHandCursor))
        self.editStartTime.setObjectName("editStartTime")
        self.gridLayout.addWidget(self.editStartTime, 1, 0, 1, 1)
        self.editEndTime = QtWidgets.QDateTimeEdit(self.groupBox)
        self.editEndTime.setCursor(QtGui.QCursor(QtCore.Qt.PointingHandCursor))
        self.editEndTime.setObjectName("editEndTime")
        self.gridLayout.addWidget(self.editEndTime, 1, 1, 1, 1)
        self.checkBoxStart = QtWidgets.QCheckBox(self.groupBox)
        self.checkBoxStart.setCursor(QtGui.QCursor(QtCore.Qt.PointingHandCursor))
        self.checkBoxStart.setChecked(False)
        self.checkBoxStart.setObjectName("checkBoxStart")
        self.gridLayout.addWidget(self.checkBoxStart, 0, 0, 1, 1)
        self.checkBoxEnd = QtWidgets.QCheckBox(self.groupBox)
        self.checkBoxEnd.setCursor(QtGui.QCursor(QtCore.Qt.PointingHandCursor))
        self.checkBoxEnd.setChecked(False)
        self.checkBoxEnd.setObjectName("checkBoxEnd")
        self.gridLayout.addWidget(self.checkBoxEnd, 0, 1, 1, 1)
        self.verticalLayout_2.addWidget(self.groupBox)
        self.label_message = QtWidgets.QLabel(self.setting)
        self.label_message.setObjectName("label_message")
        self.verticalLayout_2.addWidget(self.label_message)
        self.plainTextEdit = QtWidgets.QPlainTextEdit(self.setting)
        self.plainTextEdit.viewport().setProperty("cursor", QtGui.QCursor(QtCore.Qt.IBeamCursor))
        self.plainTextEdit.setObjectName("plainTextEdit")
        self.verticalLayout_2.addWidget(self.plainTextEdit)
        self.tabWidget.addTab(self.setting, "")
        self.verticalLayout_3.addWidget(self.tabWidget)
        MainWindow.setCentralWidget(self.centralwidget)
        self.statusbar = QtWidgets.QStatusBar(MainWindow)
        self.statusbar.setObjectName("statusbar")
        self.verticalLayout_3.addWidget(self.tabWidget)
        MainWindow.setStatusBar(self.statusbar)
        self.tabWidget.setCurrentIndex(0)
        QtCore.QMetaObject.connectSlotsByName(MainWindow)
        self.retranslateUi(MainWindow)

    def retranslateUi(self, MainWindow):
        _translate = QtCore.QCoreApplication.translate
        MainWindow.setWindowTitle(_translate("MainWindow", "Tool"))
        self.label.setText(_translate("MainWindow", "Account Ticket"))
        self.lineEditUser.setPlaceholderText(_translate("MainWindow", "Username"))
        self.lineEditPassword.setEchoMode(QtWidgets.QLineEdit.Password)
        self.lineEditPassword.setPlaceholderText(_translate("MainWindow", "Password"))
        self.browserButton.setText(_translate("MainWindow", "Browser"))
        self.tabWidget.setTabText(self.tabWidget.indexOf(self.Run), _translate("MainWindow", "Run"))
        self.editStartTime.setDisplayFormat(_translate("MainWindow", "dd/MM/yyyy HH:mm"))
        self.checkBoxStart.setText(_translate("MainWindow", "Start time"))
        self.checkBoxEnd.setText(_translate("MainWindow", "End Time"))
        self.label_message.setText(_translate("MainWindow", "Messange reply"))
        self.tabWidget.setTabText(self.tabWidget.indexOf(self.setting),
                                  _translate("MainWindow", "Setting"))
        self.buttonBox.button(QtWidgets.QDialogButtonBox.Cancel).setEnabled(False)
        self.buttonBox.button(QtWidgets.QDialogButtonBox.Ok).setText('Run')
        self.buttonBox.button(QtWidgets.QDialogButtonBox.Cancel).setText('Stop')
        item = self.listWidget.item(0)
        item.setText(_translate("Form", "History"))
        self.plainTextEdit.setPlainText("Dear anh/chị,\nSCC nhận thông tin.\n")
        self.tabWidget.setTabText(self.tabWidget.indexOf(self.setting),
                                  _translate("MainWindow", "Setting"))


class GUI(QtWidgets.QMainWindow, Ui_MainWindow):
    def __init__(self, mainwindow):
        """ Lấy ui từ Ui_MainWindow và khởi tạo attrs."""
        QtWidgets.QMainWindow.__init__(self)
        self.setupUi(mainwindow)
        self.staticfile = 'static.json'
        # get data from static file
        data = utils.read(self.staticfile)
        if data:
            self.username = data['username']
            self.password = data['password']
            self.browser = data['browser']
            self.message = data['message']
            # load static element
            self.lineEditUser.setText(self.username)
            self.lineEditPassword.setText(self.password)
            self.plainTextEdit.setPlainText(self.message)
        else:
            self.browser = r'C:\Program Files\Mozilla Firefox\firefox.exe'
        # load static element
        now = time.localtime()
        dateTime = (now.tm_year, now.tm_mon, now.tm_mday)
        self.editStartTime.setTime(QtCore.QTime(7, 0))
        self.editStartTime.setDate(QtCore.QDate(*dateTime))
        self.editEndTime.setTime(QtCore.QTime(17, 30))
        self.editEndTime.setDate(QtCore.QDate(*dateTime))
        # connect function
        self.browserButton.clicked.connect(self.get_browser)
        self.buttonBox.accepted.connect(self.accept)
        self.buttonBox.rejected.connect(self.reject)
        self.checkBoxStart.clicked.connect(lambda: self.checkBox(startfunc=True))
        self.checkBoxEnd.clicked.connect(lambda: self.checkBox(endfunc=True))
        self.plainTextEdit.textChanged.connect(self.get_plain_text)
        self.lineEditUser.editingFinished.connect(lambda: self.account_changed(changeUsername=True))
        self.lineEditPassword.editingFinished.connect(
            lambda: self.account_changed(changePassword=True))
        self.editStartTime.dateTimeChanged.connect(
            lambda: self.datetime_changed(changeStartTime=True))
        self.editEndTime.dateTimeChanged.connect(lambda: self.datetime_changed(changeEndTime=True))

    def get_browser(self):
        """ Set path firefox broser."""
        self.browser, _ = QtWidgets.QFileDialog.getOpenFileName(self, 'Select browser', '')

    def get_plain_text(self):
        """ Set message reply."""
        self.message = self.plainTextEdit.toPlainText()

    def account_changed(self, changeUsername=False, changePassword=False):
        """ live update account."""
        if hasattr(self, 'mainThread') and self.mainThread.isRunning():
            if changeUsername and self.username != self.lineEditUser.text():
                self.username = self.lineEditUser.text()
                self.mainThread.username = self.lineEditUser.text()
                self.onEventLog(f'{utils.onTime()} - SETTING - username changed')
            if changePassword and self.password != self.lineEditPassword.text():
                self.password = self.lineEditUser.text()
                self.mainThread.password = self.lineEditPassword.text()
                self.onEventLog(f'{utils.onTime()} - SETTING - password changed')

    def onEventLog(self, string):
        """ Ghi log vào listWidget - history."""
        # add item, scroll tới item cuối
        self.listWidget.addItem(string)
        count = self.listWidget.count() - 1
        self.listWidget.scrollToItem(self.listWidget.item(count))
        # max history 150 dòng, quá 150 xóa log đầu
        if count >= 150:
            item = self.listWidget.takeItem(1)
            item = None

    def refreshButtonBox(self):
        self.buttonBox.button(QtWidgets.QDialogButtonBox.Ok).setText('Run')
        self.buttonBox.button(QtWidgets.QDialogButtonBox.Cancel).setEnabled(False)
        self.buttonBox.button(QtWidgets.QDialogButtonBox.Ok).setEnabled(True)

    def checkBox(self, startfunc=False, endfunc=False):
        """ Run khi tick vào checkbox Time."""
        if startfunc:
            if self.checkBoxStart.isChecked():
                self.onEventLog(f'{utils.onTime()} - SETTING - Start at {self.editStartTime.text()}')
                self.timeThread(startfunc=True)
            else:
                self.timeStartThread.terminate()
        if endfunc:
            if self.checkBoxEnd.isChecked():
                self.onEventLog(f'{utils.onTime()} - SETTING - End at {self.editEndTime.text()}')
                self.timeThread(endfunc=True)
            else:
                self.timeEndThread.terminate()

    def timeThread(self, startfunc=False, endfunc=False):
        """ Run thread calc time."""
        if startfunc:
            self.timeStartThread = RunTimeThread(self.editStartTime.dateTime())
            self.timeStartThread.signal.connect(lambda: self.calcTimeDone(startfunc=True))
            self.timeStartThread.start()
        if endfunc:
            self.timeEndThread = RunTimeThread(self.editEndTime.dateTime())
            self.timeEndThread.signal.connect(lambda: self.calcTimeDone(endfunc=True))
            self.timeEndThread.start()

    def calcTimeDone(self, startfunc=False, endfunc=False):
        if startfunc:
            if not (hasattr(self, 'mainThread') and self.mainThread.isRunning()):
                self.tabWidget.setCurrentIndex(0)
                self.accept()
            self.checkBoxStart.setChecked(False)
            self.timeStartThread.terminate()
        if endfunc:
            if hasattr(self, 'mainThread') and self.mainThread.isRunning():
                self.tabWidget.setCurrentIndex(0)
                self.reject(force=True)
            self.checkBoxEnd.setChecked(False)
            self.timeEndThread.terminate()

    def datetime_changed(self, changeStartTime=False, changeEndTime=False):
        """ Run when change datatime."""
        if changeStartTime:
            if self.checkBoxStart.isChecked():
                self.timeStartThread.timeset = self.editStartTime.dateTime()
                self.onEventLog(f'{utils.onTime()} - SETTING - Start at {self.editStartTime.text()}')
        if changeEndTime:
            if self.checkBoxEnd.isChecked():
                self.timeEndThread.timeset = self.editEndTime.dateTime()
                self.onEventLog(f'{utils.onTime()} - SETTING - End at {self.editEndTime.text()}')

    def accept(self):
        """ Start process."""
        self.username = self.lineEditUser.text()
        self.password = self.lineEditPassword.text()
        self.message = self.plainTextEdit.toPlainText()
        if self.username.replace(' ', '') == '' or self.password.replace(' ', '') == '':
            QtWidgets.QMessageBox.information(self, "notify!", "Missing input!!")
            return
        self.buttonBox.button(QtWidgets.QDialogButtonBox.Cancel).setEnabled(True)
        self.buttonBox.button(QtWidgets.QDialogButtonBox.Ok).setEnabled(False)
        self.buttonBox.button(QtWidgets.QDialogButtonBox.Ok).setText('Running')
        # run thread main
        self.mainThread = RunMainThread(self.username, self.password, self.browser, self.message)
        self.mainThread.eventLog.connect(self.onEventLog)
        self.mainThread.start()
        # save info into staticfile
        data = {'username': self.username, 'password': self.password,
                'message': self.message, 'browser': self.browser}
        utils.write(self.staticfile, data)

    def reject(self, force=False):
        """ Hủy progress đang chạy.
            force = True  - bỏ qua bước notify
        """
        if not force:
            choice = QtWidgets.QMessageBox.question(self, 'Notify!', "Stop?",
                                                    QtWidgets.QMessageBox.Yes | QtWidgets.QMessageBox.No)
            if choice == QtWidgets.QMessageBox.No:
                return
        self.mainThread.terminate()
        self.close()
        self.refreshButtonBox()


class RunMainThread(QtCore.QThread):
    """Run a counter thread."""
    eventLog = QtCore.pyqtSignal('QString')

    def __init__(self, username, password, browser, message):
        super().__init__()
        self.username = username
        self.password = password
        self.browser = browser
        self.message = message

    def __del__(self):
        self.wait()

    def terminate(self):
        self.eventLog.emit(f'{utils.onTime()} - INFO - Stop')
        super().terminate()

    def run(self):
        self.eventLog.emit(f'{utils.onTime()} - INFO - Running')
        # Main()
        main(self.username, self.password, self.browser, self.message, self.eventLog)


class RunTimeThread(QtCore.QThread):
    signal = QtCore.pyqtSignal()

    def __init__(self, timeset):
        super().__init__()
        self.timeset = timeset

    def __del__(self):
        self.wait()

    def run(self):
        while True:
            now = time.localtime()
            dateTime = (now.tm_year, now.tm_mon, now.tm_mday, now.tm_hour, now.tm_min)
            if self.timeset <= QtCore.QDateTime(*dateTime):
                self.signal.emit()
            time.sleep(30)


def get_logging():
    """ Setup global logging."""
    global logger
    # create log folder
    if not os.path.isdir('./log'):
        os.mkdir('./log')
    logFormatter = logging.Formatter("%(asctime)s - %(levelname)s - %(message)s")
    logger = logging.getLogger()
    logger.setLevel(logging.INFO)
    # file log
    fileHandler = logging.FileHandler(f'log/{time.strftime("%Y-%m-%d", time.gmtime())}.log')
    fileHandler.setFormatter(logFormatter)
    logger.addHandler(fileHandler)
    # console log
    consoleHandler = logging.StreamHandler()
    consoleHandler.setFormatter(logFormatter)
    logger.addHandler(consoleHandler)


if __name__ == '__main__':
    get_logging()

    app = QtWidgets.QApplication(sys.argv)
    app.setStyle('Fusion')

    MainWindow = QtWidgets.QMainWindow()
    ui = GUI(MainWindow)
    MainWindow.show()
    sys.exit(app.exec_())

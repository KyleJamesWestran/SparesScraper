import os
from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtWidgets import QMessageBox, QFileDialog
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from selenium.common.exceptions import NoSuchElementException, TimeoutException
from selenium.webdriver.chrome.service import Service
from subprocess import CREATE_NO_WINDOW

class Ui_Form(object):
    def setupUi(self, Form):
        # UI Setup
        Form.setObjectName("Form")
        Form.resize(850, 380)
        Form.setWindowIcon(QtGui.QIcon('Resources/icon.png'))
        Form.setMinimumSize(QtCore.QSize(850, 380))
        self.gridLayout = QtWidgets.QGridLayout(Form)
        self.gridLayout.setObjectName("gridLayout")

        self.lblPart = QtWidgets.QLabel(Form)
        self.lblPart.setObjectName("lblPart")
        self.gridLayout.addWidget(self.lblPart, 2, 1, 1, 1)

        self.label_3 = QtWidgets.QLabel(Form)
        self.label_3.setObjectName("label_3")
        self.gridLayout.addWidget(self.label_3, 7, 0, 1, 1)

        self.lblOutput = QtWidgets.QLabel(Form)
        self.lblOutput.setObjectName("lblOutput")
        self.gridLayout.addWidget(self.lblOutput, 3, 1, 1, 1)

        self.label_4 = QtWidgets.QLabel(Form)
        self.label_4.setMaximumSize(QtCore.QSize(180, 16777215))
        self.label_4.setObjectName("label_4")
        self.gridLayout.addWidget(self.label_4, 3, 0, 1, 1)

        self.label = QtWidgets.QLabel(Form)
        self.label.setMinimumSize(QtCore.QSize(0, 50))
        self.label.setMaximumSize(QtCore.QSize(16777215, 50))
        self.label.setObjectName("label")
        self.gridLayout.addWidget(self.label, 0, 0, 1, 4)

        self.label_5 = QtWidgets.QLabel(Form)
        self.label_5.setMaximumSize(QtCore.QSize(180, 16777215))
        self.label_5.setObjectName("label_5")
        self.gridLayout.addWidget(self.label_5, 1, 0, 1, 1)

        self.input = QtWidgets.QLineEdit(Form)
        self.label_5.setObjectName("label_5")
        self.gridLayout.addWidget(self.input, 1, 1, 1, 3)

        self.label_6 = QtWidgets.QLabel(Form)
        self.label_6.setMaximumSize(QtCore.QSize(180, 16777215))
        self.label_6.setObjectName("label_6")
        self.gridLayout.addWidget(self.label_6, 6, 0, 1, 1)

        self.checkProcess = QtWidgets.QCheckBox(Form)
        self.label_5.setObjectName("checkProcess")
        self.gridLayout.addWidget(self.checkProcess, 6, 1, 1, 1)

        self.btnProcess = QtWidgets.QPushButton(Form)
        self.btnProcess.setMinimumSize(QtCore.QSize(0, 30))
        self.btnProcess.setMaximumSize(QtCore.QSize(120, 60))
        icon = QtGui.QIcon()
        icon.addPixmap(QtGui.QPixmap("Resources/refresh.png"), QtGui.QIcon.Normal, QtGui.QIcon.Off)
        self.btnProcess.setIcon(icon)
        self.btnProcess.setIconSize(QtCore.QSize(15, 15))
        self.btnProcess.setObjectName("btnProcess")
        self.gridLayout.addWidget(self.btnProcess, 7, 3, 1, 1)

        self.btnOutputBrowse = QtWidgets.QPushButton(Form)
        self.btnOutputBrowse.setMinimumSize(QtCore.QSize(0, 30))
        self.btnOutputBrowse.setMaximumSize(QtCore.QSize(120, 60))
        icon1 = QtGui.QIcon()
        icon1.addPixmap(QtGui.QPixmap("Resources/search.png"), QtGui.QIcon.Normal, QtGui.QIcon.Off)
        self.btnOutputBrowse.setIcon(icon1)
        self.btnOutputBrowse.setIconSize(QtCore.QSize(13, 13))
        self.btnOutputBrowse.setObjectName("btnOutputBrowse")
        self.gridLayout.addWidget(self.btnOutputBrowse, 3, 3, 1, 1)

        self.groupSites = QtWidgets.QGroupBox(Form)
        self.groupSites.setObjectName("groupSites")
        self.gridLayout_2 = QtWidgets.QGridLayout(self.groupSites)
        self.gridLayout_2.setObjectName("gridLayout_2")
        self.chkNationwide = QtWidgets.QCheckBox(self.groupSites)
        self.chkNationwide.setObjectName("chkNationwide")
        self.gridLayout_2.addWidget(self.chkNationwide, 1, 1, 1, 1)
        self.chkElectrolux = QtWidgets.QCheckBox(self.groupSites)
        self.chkElectrolux.setObjectName("chkElectrolux")
        self.gridLayout_2.addWidget(self.chkElectrolux, 0, 0, 1, 1)
        self.chkStokesAP = QtWidgets.QCheckBox(self.groupSites)
        self.chkStokesAP.setObjectName("chkStokesAP")
        self.gridLayout_2.addWidget(self.chkStokesAP, 2, 3, 1, 1)
        self.chkuBuy = QtWidgets.QCheckBox(self.groupSites)
        self.chkuBuy.setObjectName("chkuBuy")
        self.gridLayout_2.addWidget(self.chkuBuy, 2, 1, 1, 1)
        self.chkStateAS = QtWidgets.QCheckBox(self.groupSites)
        self.chkStateAS.setObjectName("chkStateAS")
        self.gridLayout_2.addWidget(self.chkStateAS, 2, 2, 1, 1)
        self.chkOvenPA = QtWidgets.QCheckBox(self.groupSites)
        self.chkOvenPA.setObjectName("chkOvenPA")
        self.gridLayout_2.addWidget(self.chkOvenPA, 3, 1, 1, 1)
        self.chkOnlineAP = QtWidgets.QCheckBox(self.groupSites)
        self.chkOnlineAP.setObjectName("chkOnlineAP")
        self.gridLayout_2.addWidget(self.chkOnlineAP, 1, 0, 1, 1)
        self.chkWholesale = QtWidgets.QCheckBox(self.groupSites)
        self.chkWholesale.setObjectName("chkWholesale")
        self.gridLayout_2.addWidget(self.chkWholesale, 1, 2, 1, 1)
        self.chkEllisE = QtWidgets.QCheckBox(self.groupSites)
        self.chkEllisE.setObjectName("chkEllisE")
        self.gridLayout_2.addWidget(self.chkEllisE, 3, 3, 1, 1)
        self.chkOnlineAS = QtWidgets.QCheckBox(self.groupSites)
        self.chkOnlineAS.setObjectName("chkOnlineAS")
        self.gridLayout_2.addWidget(self.chkOnlineAS, 1, 3, 1, 1)
        self.chkDoug = QtWidgets.QCheckBox(self.groupSites)
        self.chkDoug.setObjectName("chkDoug")
        self.gridLayout_2.addWidget(self.chkDoug, 0, 2, 1, 1)
        self.chkFisher = QtWidgets.QCheckBox(self.groupSites)
        self.chkFisher.setObjectName("chkFisher")
        self.gridLayout_2.addWidget(self.chkFisher, 0, 1, 1, 1)
        self.chkMrAp = QtWidgets.QCheckBox(self.groupSites)
        self.chkMrAp.setObjectName("chkMrAp")
        self.gridLayout_2.addWidget(self.chkMrAp, 0, 3, 1, 1)
        self.chkGenuineAP = QtWidgets.QCheckBox(self.groupSites)
        self.chkGenuineAP.setObjectName("chkGenuineAP")
        self.gridLayout_2.addWidget(self.chkGenuineAP, 2, 0, 1, 1)
        self.chkStoveC = QtWidgets.QCheckBox(self.groupSites)
        self.chkStoveC.setObjectName("chkStoveC")
        self.gridLayout_2.addWidget(self.chkStoveC, 3, 2, 1, 1)
        self.chkDiscountAP = QtWidgets.QCheckBox(self.groupSites)
        self.chkDiscountAP.setObjectName("chkDiscountAP")
        self.gridLayout_2.addWidget(self.chkDiscountAP, 3, 0, 1, 1)
        self.chkSydneyAS = QtWidgets.QCheckBox(self.groupSites)
        self.chkSydneyAS.setObjectName("chkSydneyAS")
        self.gridLayout_2.addWidget(self.chkSydneyAS, 5, 0, 1, 1)

        self.btnCheck = QtWidgets.QPushButton(self.groupSites)
        icon2 = QtGui.QIcon()
        icon2.addPixmap(QtGui.QPixmap("Resources/select.png"), QtGui.QIcon.Normal, QtGui.QIcon.Off)
        self.btnCheck.setIcon(icon2)
        self.btnCheck.setIconSize(QtCore.QSize(12, 12))
        self.btnCheck.setObjectName("btnCheck")
        self.gridLayout_2.addWidget(self.btnCheck, 5, 2, 1, 1)

        self.btnUncheck = QtWidgets.QPushButton(self.groupSites)
        icon3 = QtGui.QIcon()
        icon3.addPixmap(QtGui.QPixmap("Resources/cancel.png"), QtGui.QIcon.Normal, QtGui.QIcon.Off)
        self.btnUncheck.setIcon(icon3)
        self.btnUncheck.setIconSize(QtCore.QSize(12, 12))
        self.btnUncheck.setObjectName("btnUncheck")
        self.gridLayout_2.addWidget(self.btnUncheck, 5, 3, 1, 1)
        self.gridLayout.addWidget(self.groupSites, 4, 0, 1, 4)

        self.btnListBrowse = QtWidgets.QPushButton(Form)
        self.btnListBrowse.setMinimumSize(QtCore.QSize(0, 30))
        self.btnListBrowse.setMaximumSize(QtCore.QSize(120, 60))
        self.btnListBrowse.setIcon(icon1)
        self.btnListBrowse.setIconSize(QtCore.QSize(13, 13))
        self.btnListBrowse.setObjectName("btnListBrowse")
        self.gridLayout.addWidget(self.btnListBrowse, 2, 3, 1, 1)

        self.label_2 = QtWidgets.QLabel(Form)
        self.label_2.setMaximumSize(QtCore.QSize(180, 16777215))
        self.label_2.setObjectName("label_2")
        self.gridLayout.addWidget(self.label_2, 2, 0, 1, 1)

        spacerItem = QtWidgets.QSpacerItem(50, 50, QtWidgets.QSizePolicy.Minimum, QtWidgets.QSizePolicy.Expanding)
        self.gridLayout.addItem(spacerItem, 5, 0, 1, 4)

        self.lblStatus = QtWidgets.QLabel(Form)
        self.lblStatus.setObjectName("lblStatus")
        self.gridLayout.addWidget(self.lblStatus, 7, 1, 1, 2)

        self.retranslateUi(Form)
        QtCore.QMetaObject.connectSlotsByName(Form)

        # Function Calling
        self.input.textChanged.connect(self.disable)
        self.btnListBrowse.clicked.connect(lambda: selectInput(self))
        self.btnOutputBrowse.clicked.connect(lambda: selectOutput(self))
        self.btnUncheck.clicked.connect(lambda: uncheckAll(self))
        self.btnCheck.clicked.connect(lambda: checkAll(self))
        self.btnProcess.clicked.connect(lambda: inputCheck(self))

    def retranslateUi(self, Form):
        _translate = QtCore.QCoreApplication.translate
        Form.setWindowTitle(_translate("Form", "Spares Scraper"))
        self.lblPart.setText(_translate("Form", "Press browse"))
        self.label_3.setText(_translate("Form", "Program status:"))
        self.label_6.setText(_translate("Form", "View Process:"))
        self.lblOutput.setText(_translate("Form", "Press browse"))
        self.label_4.setText(_translate("Form", "Select save location:"))
        self.label_5.setText(_translate("Form", "Insert Part or browse below:"))
        self.label.setText(_translate("Form", "<html><head/><body><p align=\"center\"><span style=\" font-size:28pt; font-weight:600; color:#0f3362;\">SPARES SCRAPER</span></p></body></html>"))
        self.btnProcess.setText(_translate("Form", " Process"))
        self.btnOutputBrowse.setText(_translate("Form", " Browse"))
        self.groupSites.setTitle(_translate("Form", "Select Sites to Scrape"))
        self.chkNationwide.setText(_translate("Form", "Nationwide Spares"))
        self.chkElectrolux.setText(_translate("Form", "Electrolux"))
        self.chkStokesAP.setText(_translate("Form", "Stokes Appliance Parts"))
        self.chkuBuy.setText(_translate("Form", "uBuyOz"))
        self.chkStateAS.setText(_translate("Form", "Statewide Appliance Spares"))
        self.chkOvenPA.setText(_translate("Form", "Oven Parts Australia"))
        self.chkOnlineAP.setText(_translate("Form", "Online Appliance Parts"))
        self.chkWholesale.setText(_translate("Form", "Wholesale Appliance Supplies"))
        self.chkEllisE.setText(_translate("Form", "Ellis Electrical"))
        self.chkOnlineAS.setText(_translate("Form", "Online Appliance Spares"))
        self.chkDoug.setText(_translate("Form", "Doug Smith Spares"))
        self.chkFisher.setText(_translate("Form", "Fisher and Paykel"))
        self.chkMrAp.setText(_translate("Form", "Mr Appliance"))
        self.chkGenuineAP.setText(_translate("Form", "Genuine Appliance Parts"))
        self.chkStoveC.setText(_translate("Form", "Stove Connection"))
        self.chkDiscountAP.setText(_translate("Form", "Discount Appliance Parts"))
        self.chkSydneyAS.setText(_translate("Form", "Sydney Appliance Service"))
        self.btnCheck.setText(_translate("Form", " Check All"))
        self.btnUncheck.setText(_translate("Form", " Uncheck All"))
        self.btnListBrowse.setText(_translate("Form", " Browse"))
        self.label_2.setText(_translate("Form", "Select a parts list:"))
        self.lblStatus.setText(_translate("Form", "<html><head/><body><p><span style=\" color:#0f3362;\">Waiting for process to begin</span></p></body></html>"))

    def disable(self):
        if self.input.text() == "":
            self.btnListBrowse.setEnabled(True)
        else:
            self.btnListBrowse.setEnabled(False)

# Functions
def selectInput(self):
    default = self.label.styleSheet()
    self.lblPart.setStyleSheet(default)
    path = str(QFileDialog.getOpenFileNames(None, "select File","c:/","Excel files(*.xlsx)"))
    path = str(path)
    path = path.replace("(['","")
    path = path.replace("'], 'Excel files(*.xlsx)')", "")


    if path == "([], '')":
        self.lblPart.setText("Press browse")
        self.input.setEnabled(True)
    else:
        global parts
        self.lblPart.setText(path)
        data = pd.read_excel(path, header = None)
        data.dropna(0, inplace = True)
        data.rename( columns={0:'Parts'}, inplace=True )
        parts = data['Parts'].tolist()
        self.input.setEnabled(False)

def selectOutput(self):
    default = self.label.styleSheet()
    self.lblOutput.setStyleSheet(default)
    path = str(QFileDialog.getExistingDirectory(None, "select Directory","c://"))

    if path == "":
        self.lblOutput.setText("Press browse")
    else:
        self.lblOutput.setText(path)

def uncheckAll(self):
    for item in self.groupSites.findChildren(QtWidgets.QCheckBox):
        item.setChecked(False)

def checkAll(self):
    for item in self.groupSites.findChildren(QtWidgets.QCheckBox):
        item.setChecked(True)

def inputCheck(self):
    writeLine(self,"This is about to get removed!")
    os.remove("tempData.csv")
    errors = False
    default = self.label_2.styleSheet()

    self.lblPart.setStyleSheet(default)
    self.lblOutput.setStyleSheet(default)

    if self.input.text() == "":
        if self.lblPart.text() == "Press browse":
            self.lblPart.setStyleSheet("color: red")
            errors = True
    else:
        errors = False

    if self.lblOutput.text() == "Press browse":
        self.lblOutput.setStyleSheet("color: red")
        errors = True

    if errors == True:
        fixErrorMessage()
    else:
        self.lblPart.setStyleSheet(default)
        self.lblOutput.setStyleSheet(default)
        self.lblStatus.setText("Processing Request, Please Wait...")
        self.lblStatus.setStyleSheet("color: orange")
        QtWidgets.QApplication.processEvents()

        process(self)

def process(self):
    if self.checkProcess.isChecked() == False:
        options.add_argument("--headless")

    writeLine(self,"Part Number,Alternative Number,Price,Discounted Price")

    if self.input.text() != "":
        global parts
        input = self.input.text()
        parts = [input]

    #Disable User Input
    for item in self.groupSites.findChildren(QtWidgets.QCheckBox):
        item.setEnabled(False)
    self.btnListBrowse.setEnabled(False)
    self.btnOutputBrowse.setEnabled(False)
    self.btnCheck.setEnabled(False)
    self.btnUncheck.setEnabled(False)
    self.btnProcess.setEnabled(False)
    self.checkProcess.setEnabled(False)
    self.input.setEnabled(False)
    driver = webdriver.Chrome(options=options, service=service)
    driver.maximize_window()

    # If statements to scrape each website
    if self.chkElectrolux.isChecked() == True:
        electrolux(self,driver)

    if self.chkFisher.isChecked() == True:
        fisher(self,driver)

    if self.chkDoug.isChecked() == True:
        doug(self,driver)

    if self.chkMrAp.isChecked() == True:
        mrApp(self,driver)

    if self.chkOnlineAP.isChecked() == True:
        onlineAP(self,driver)

    if self.chkNationwide.isChecked() == True:
        nation(self,driver)

    if self.chkWholesale.isChecked() == True:
        wholesale(self,driver)

    if self.chkOnlineAS.isChecked() == True:
        onlineAS(self,driver)

    if self.chkGenuineAP.isChecked() == True:
        genuineAP(self,driver)

    if self.chkuBuy.isChecked() == True:
        uBuy(self,driver)

    if self.chkStateAS.isChecked() == True:
        stateAS(self,driver)

    if self.chkStokesAP.isChecked() == True:
        stokesAP(self,driver)

    if self.chkDiscountAP.isChecked() == True:
        discountAP(self,driver)

    if self.chkOvenPA.isChecked() == True:
        ovenAP(self,driver)

    if self.chkStoveC.isChecked() == True:
        stoveC(self,driver)

    if self.chkEllisE.isChecked() == True:
        ellis(self,driver)

    if self.chkSydneyAS.isChecked() == True:
        sydneyAS(self,driver)

    driver.close()
    completeStatus(self)

# Scraping Functions & Options
service = Service("Resources/chromedriver.exe")
service.creationflags = CREATE_NO_WINDOW
options = webdriver.ChromeOptions()
options.add_argument('--window-size=1920,1080')

def electrolux(self,driver):
    writeLine(self,"Electrolux")
    print(parts)
    try:
        # Data to scrape from website
        for part in parts:
            # User Aid
            length = len(parts)
            position = (parts.index(part))+1
            self.lblStatus.setText("Scraping Electrolux {0} of {1}".format(position,length))
            self.lblStatus.setStyleSheet("color: #FF6E00")
            QtWidgets.QApplication.processEvents()

            driver.get("https://shop.electrolux.com.au/")
            driver.implicitly_wait(3)

            #Insert Part Number
            driver.find_element(By.XPATH, "/html/body/header/div/nav/div[2]/div[1]/input[1]").send_keys(part)

            #Search
            driver.find_element(By.XPATH, "/html/body/header/div/nav/div[2]/div[1]/input[1]").send_keys(u'\ue007')
            driver.implicitly_wait(2)

            #Try Except!
            try:
                # Get Price
                price = driver.find_element(By.XPATH, "/html/body/div[2]/div/div[2]/div[3]/div/div/ul/li/div/div[1]/div[2]/div/div/span").text

                # Get Alternative Part Number
                alternative = driver.find_element(By.XPATH, "/html/body/div[2]/div/div[2]/div[3]/div/div/ul/li/div/div[1]/div[2]/a[2]").text

                # Get Discounted
                discounted = "None"

                partInfo = ("{0}, {1}, {2}, {3}".format(part, alternative, price, discounted))

            except NoSuchElementException:
                partInfo = "{0}, None, None, None".format(part)

            writeLine(self,partInfo)

    except TimeoutException:
        writeLine(self,"Unable,to,access,website")

def fisher(self,driver):
    writeLine(self,"Fisher & Paykel")
    try:
        # Data to scrape from website
        for part in parts:
            # User Aid
            length = len(parts)
            position = (parts.index(part)) + 1
            self.lblStatus.setText("Scraping Fisher & Paykel {0} of {1}".format(position, length))
            self.lblStatus.setStyleSheet("color: #FF6E00")
            QtWidgets.QApplication.processEvents()

            driver.get("https://www.fisherpaykel.com/au/spare-parts/")
            driver.implicitly_wait(10)

            # Press search button
            driver.find_element(By.XPATH, "/html/body/div[3]/div[1]/header/nav/div[1]/div/div[1]/div/div[1]/button").click()

            # Insert Part Number
            driver.find_element(By.XPATH, "/html/body/div[3]/div[1]/header/nav/div[1]/div/div[2]/div[1]/div[1]/form/input[1]").send_keys(part)

            # Search
            driver.find_element(By.XPATH, "/html/body/div[3]/div[1]/header/nav/div[1]/div/div[2]/div[1]/div[1]/form/input[1]").send_keys(u'\ue007')

            # Try Except!
            try:
                # Get Price
                price = driver.find_element(By.XPATH, "/html/body/div[3]/div[4]/div[5]/div/div[1]/div[2]/div[2]/div/div[2]/div/div/div[2]/div[7]/div/span/span/span").text

                # Get Alternative Part Number
                alternative = "None"

                # Get Discounted
                discounted = "None"

                partInfo = ("{0}, {1}, {2}, {3}".format(part, alternative, price, discounted))

            except NoSuchElementException:
                partInfo = "{0}, None, None, None".format(part)

            # Send To csv
            writeLine(self,partInfo)

    except TimeoutException:
        writeLine(self,"Unable,to,access,website")

def doug(self,driver):
    writeLine(self,"Doug Smith Spares")
    try:
        # Data to scrape from website
        for part in parts:
            # User Aid
            length = len(parts)
            position = (parts.index(part)) + 1
            self.lblStatus.setText("Scraping Doug Smith Spares {0} of {1}".format(position, length))
            self.lblStatus.setStyleSheet("color: #FF6E00")
            QtWidgets.QApplication.processEvents()

            driver.get("https://dougsmithspares.com.au/shop/")

            # Press search button
            button = driver.find_element(By.XPATH, "//*[@id='et_search_icon']")
            ActionChains(driver).move_to_element(button).click(button)

            # Insert Part Number
            driver.find_element(By.XPATH, "/html/body/div[1]/header/div[2]/div/form/input").send_keys(part)

            # Search
            driver.find_element(By.XPATH, "/html/body/div[1]/header/div[2]/div/form/input").send_keys(u'\ue007')

            # Press item
            driver.find_element(By.XPATH, "//*[@id='left-area']").click()

            # Try Except!
            try:
                # Get Price
                price = driver.find_element(By.XPATH, "/html/body/div[1]/div/div/div/div/div[1]/div[2]/div[1]/div[2]/p[1]/del/span/bdi").text

                # Get Discounted Price
                discounted = driver.find_element(By.XPATH, "/html/body/div[1]/div/div/div/div/div[1]/div[2]/div[1]/div[2]/p[1]/ins/span/bdi").text

                # Get Alternative Part Number
                alternative = "None"

                partInfo = ("{0}, {1}, {2}, {3}".format(part, alternative, price, discounted))

            except NoSuchElementException:
                try:
                    # Get Price
                    price = driver.find_element(By.XPATH, "/html/body/div[1]/div/div/div/div/div[1]/div[2]/div[1]/div[2]/p[1]/span/bdi").text

                    # Get Alternative Part Number
                    alternative = "None"

                    # Get Discounted Price
                    discounted = "None"

                    partInfo = "{0}, {1}, {2}, {3}".format(part, alternative, price, discounted)

                except NoSuchElementException:
                    partInfo = "{0}, None, None, None".format(part)

            # Send To csv
            writeLine(self,partInfo)

    except TimeoutException:
        writeLine(self,"Unable,to,access,website")

def mrApp(self,driver):
    writeLine(self,"Mr Appliance")
    try:
        # Data to scrape from website
        for part in parts:
            # User Aid
            length = len(parts)
            position = (parts.index(part)) + 1
            self.lblStatus.setText("Scraping Mr Appliance {0} of {1}".format(position, length))
            self.lblStatus.setStyleSheet("color: #FF6E00")
            QtWidgets.QApplication.processEvents()

            driver.get("https://mrappliance.com.au/")

            # Insert Part Number
            driver.find_element(By.XPATH, "/html/body/div[2]/div[1]/div/div/div[1]/div/div[2]/div[1]/div/div[2]/div/div[1]/section/form/input[1]").send_keys(part)

            # Search
            driver.find_element(By.XPATH, "/html/body/div[2]/div[1]/div/div/div[1]/div/div[2]/div[1]/div/div[2]/div/div[1]/section/form/input[1]").send_keys(u'\ue007')

            # Try Except!
            try:
                # Get Price
                price = driver.find_element(By.XPATH, "/html/body/div[1]/div/div[1]/div/div[2]/div/div[2]/div[2]/p[1]/del/span/bdi").text

                # Get Discounted Price
                discounted = driver.find_element(By.XPATH, "/html/body/div[1]/div/div[1]/div/div[2]/div/div[2]/div[2]/p[1]/ins/span/bdi").text

                # Get Alternative Part Number
                alternative = "None"

                partInfo = ("{0}, {1}, {2}, {3}".format(part, alternative, price, discounted))

            except NoSuchElementException:
                try:
                    # Get Price
                    price = driver.find_element(By.XPATH, "/html/body/div[1]/div/div[1]/div/div[2]/div/div[2]/div[2]/p[1]/span/bdi").text

                    # Get Alternative Part Number
                    alternative = "None"

                    # Get Discounted Price
                    discounted = "None"

                    partInfo = "{0}, {1}, {2}, {3}".format(part, alternative, price, discounted)

                except NoSuchElementException:
                    partInfo = "{0}, None, None, None".format(part)

            # Send To csv
            writeLine(self,partInfo)

    except TimeoutException:
        writeLine(self,"Unable,to,access,website")

def onlineAP(self,driver):
    writeLine(self, "Online Appliance Parts")
    try:
        # Data to scrape from website
        for part in parts:
            # User Aid
            length = len(parts)
            position = (parts.index(part)) + 1
            self.lblStatus.setText("Scraping Online Appliance Parts {0} of {1}".format(position, length))
            self.lblStatus.setStyleSheet("color: #FF6E00")
            QtWidgets.QApplication.processEvents()

            driver.get("https://www.onlineapplianceparts.com.au/")
            driver.implicitly_wait(3)

            # Insert Part Number
            driver.find_element(By.XPATH, "//*[@id='search']/input").send_keys(part)

            # Search
            driver.find_element(By.XPATH, "//*[@id='search']/input").send_keys(u'\ue007')

            # Try Except!
            try:
                # Get Price
                price = driver.find_element(By.XPATH, "//*[@id='content-container']/div[3]/div/div/div[2]/p[2]").text
                price = price.replace(",","")
                price = price.replace("\n", " ")
                # Get Discounted Price
                discounted = "None"

                # Get Alternative Part Number
                alternative = "None"

                partInfo = ("{0}, {1}, {2}, {3}".format(part, alternative, price, discounted))

            except NoSuchElementException:
                partInfo = "{0}, None, None, None".format(part)

            # Send To csv
            writeLine(self, partInfo)

    except TimeoutException:
        writeLine(self, "Unable,to,access,website")

def nation(self,driver):
    writeLine(self, "Nationwide Spares")
    try:
        # Data to scrape from website
        for part in parts:
            # User Aid
            length = len(parts)
            position = (parts.index(part)) + 1
            self.lblStatus.setText("Scraping Nationwide Spares {0} of {1}".format(position, length))
            self.lblStatus.setStyleSheet("color: #FF6E00")
            QtWidgets.QApplication.processEvents()

            driver.get("https://www.nationwidespares.com.au/")
            driver.implicitly_wait(3)

            # Insert Part Number
            driver.find_element(By.XPATH, "//*[@id='search']/input").send_keys(part)

            # Search
            driver.find_element(By.XPATH, "//*[@id='search']/input").send_keys(u'\ue007')

            # Try Except!
            try:
                # Get Price
                price = driver.find_element(By.XPATH, "//*[@id='content']/div[3]/div/div/div[2]/p[2]").text
                price = price.replace(",", "")
                price = price.replace("\n", " ")
                # Get Discounted Price
                discounted = "None"

                # Get Alternative Part Number
                alternative = "None"

                partInfo = ("{0}, {1}, {2}, {3}".format(part, alternative, price, discounted))

            except NoSuchElementException:
                partInfo = "{0}, None, None, None".format(part)

            # Send To csv
            writeLine(self, partInfo)

    except TimeoutException:
        writeLine(self, "Unable,to,access,website")

def wholesale(self,driver):
    writeLine(self, "Wholesale Appliance Supplies")
    try:
        # Data to scrape from website
        for part in parts:
            # User Aid
            length = len(parts)
            position = (parts.index(part)) + 1
            self.lblStatus.setText("Scraping Wholesale Appliance Supplies {0} of {1}".format(position, length))
            self.lblStatus.setStyleSheet("color: #FF6E00")
            QtWidgets.QApplication.processEvents()

            driver.get("https://www.appliancesupplies.com.au/")
            driver.implicitly_wait(3)

            # Insert Part Number
            driver.find_element(By.XPATH, "//*[@id='search']/form[2]/div/div/input").send_keys(part)

            # Search
            driver.find_element(By.XPATH, "//*[@id='search']/form[2]/div/div/input").send_keys(u'\ue007')

            # Try Except!
            try:
                driver.find_element(By.XPATH, "//*[@id='content']/div[2]/a/div/div[1]").click()

                try:
                    # Get Price
                    price = driver.find_element(By.XPATH, "//*[@id='content']/section[1]/div/div[2]/div/div[1]").text

                    # Get Alternative Part Number
                    alternative = driver.find_element(By.XPATH, "//*[@id='breadcrumbs']/li[3]").text

                except NoSuchElementException:
                    driver.find_element(By.XPATH, "//*[@id='content']/section/div/p/a").click()

                    # Get Price
                    price = driver.find_element(By.XPATH, "//*[@id='content']/section[1]/div/div[2]/div/div[1]").text

                    # Get Alternative Part Number
                    alternative = driver.find_element(By.XPATH, "//*[@id='breadcrumbs']/li[3]").text

                # Get Discounted Price
                discounted = "None"

                partInfo = ("{0}, {1}, {2}, {3}".format(part, alternative, price, discounted))

            except NoSuchElementException:
                partInfo = "{0}, None, None, None".format(part)

            # Send To csv
            writeLine(self, partInfo)

    except TimeoutException:
        writeLine(self, "Unable,to,access,website")

def onlineAS(self,driver):
    writeLine(self, "Online Appliance Spares")
    try:
        # Data to scrape from website
        for part in parts:
            # User Aid
            length = len(parts)
            position = (parts.index(part)) + 1
            self.lblStatus.setText("Scraping Online Appliance Spares {0} of {1}".format(position, length))
            self.lblStatus.setStyleSheet("color: #FF6E00")
            QtWidgets.QApplication.processEvents()

            driver.get("https://shop.onlineappliancespares.com.au/")
            driver.implicitly_wait(3)

            # Insert Part Number
            driver.find_element(By.XPATH, "//*[@id='quick-search']/span/input[1]").send_keys(part)

            # Search
            driver.find_element(By.XPATH, "//*[@id='quick-search']/span/input[1]").send_keys(u'\ue007')

            # Try Except!
            try:
                # Get Price
                price = driver.find_element(By.XPATH, "//*[@id='content']/div/div/div[1]/div[2]/div[3]").text

                # Get Discounted Price
                discounted = "None"

                # Get Alternative Part Number
                alternative = driver.find_element(By.XPATH, "//*[@id='content']/ul/li[3]/a").text

                partInfo = ("{0}, {1}, {2}, {3}".format(part, alternative, price, discounted))

            except NoSuchElementException:
                partInfo = "{0}, None, None, None".format(part)

            # Send To csv
            writeLine(self, partInfo)

    except TimeoutException:
        writeLine(self, "Unable,to,access,website")

def genuineAP(self,driver):
    writeLine(self, "Genuine Appliance Parts")
    try:
        # Data to scrape from website
        for part in parts:
            # User Aid
            length = len(parts)
            position = (parts.index(part)) + 1
            self.lblStatus.setText("Scraping Genuine Appliance Parts {0} of {1}".format(position, length))
            self.lblStatus.setStyleSheet("color: #FF6E00")
            QtWidgets.QApplication.processEvents()

            driver.get("https://genuineapplianceparts.com.au/")
            driver.implicitly_wait(3)

            # Insert Part Number
            driver.find_element(By.XPATH, "//*[@id='search']").send_keys(part)

            # Search
            driver.find_element(By.XPATH, "//*[@id='search']").send_keys(u'\ue007')

            # Try Except!
            try:
                # Get Price
                price = driver.find_element(By.XPATH, "/html/body/div[1]/main/div[2]/div/div[1]/div[2]/div[2]/span[1]/span/span[2]/span").text

                # Get Discounted Price
                discounted = driver.find_element(By.XPATH, "/html/body/div[1]/main/div[2]/div/div[1]/div[2]/div[2]/span[2]/span/span[2]/span").text

                # Get Alternative Part Number
                alternative = "None"

                partInfo = ("{0}, {1}, {2}, {3}".format(part, alternative, price, discounted))

            except NoSuchElementException:
                partInfo = "{0}, None, None, None".format(part)

            # Send To csv
            writeLine(self, partInfo)

    except TimeoutException:
        writeLine(self, "Unable,to,access,website")

def uBuy(self,driver):
    writeLine(self, "uBuyOz")
    try:
        # Data to scrape from website
        for part in parts:
            # User Aid
            length = len(parts)
            position = (parts.index(part)) + 1
            self.lblStatus.setText("Scraping uBuyOz {0} of {1}".format(position, length))
            self.lblStatus.setStyleSheet("color: #FF6E00")
            QtWidgets.QApplication.processEvents()

            driver.get("https://www.ubuyoz.com.au/")
            driver.implicitly_wait(3)

            # Insert Part Number
            driver.find_element(By.XPATH, "//*[@id='blog']/div/div[4]/div[2]/form/input[2]").send_keys(part)

            # Search
            driver.find_element(By.XPATH, "//*[@id='blog']/div/div[4]/div[2]/form/input[2]").send_keys(u'\ue007')

            # Try Except!
            try:
                # Get Price
                price = driver.find_element(By.XPATH, "/html/body/div[1]/div[5]/div/div/div[2]/div[2]/p[1]/del/span/bdi").text

                # Get Discounted Price
                discounted = driver.find_element(By.XPATH, "/html/body/div[1]/div[5]/div/div/div[2]/div[2]/p[1]/ins/span/bdi").text

                # Get Alternative Part Number
                alternative = "None"

                partInfo = ("{0}, {1}, {2}, {3}".format(part, alternative, price, discounted))

            except NoSuchElementException:
                partInfo = "{0}, None, None, None".format(part)

            # Send To csv
            writeLine(self, partInfo)

    except TimeoutException:
        writeLine(self, "Unable,to,access,website")

def stateAS(self,driver):
    writeLine(self, "Statewide Appliance Spares")
    try:
        # Data to scrape from website
        for part in parts:
            # User Aid
            length = len(parts)
            position = (parts.index(part)) + 1
            self.lblStatus.setText("Scraping Statewide Appliance Spares {0} of {1}".format(position, length))
            self.lblStatus.setStyleSheet("color: #FF6E00")
            QtWidgets.QApplication.processEvents()

            driver.get("https://www.statewideapp.com.au/")
            driver.implicitly_wait(3)

            # Insert Part Number
            driver.find_element(By.XPATH, "//*[@id='small-searchterms']").send_keys(part)

            # Search
            driver.find_element(By.XPATH, "//*[@id='small-searchterms']").send_keys(u'\ue007')

            # Try Except!
            try:
                # Get Price
                price = driver.find_element(By.XPATH, "/html/body/div[6]/section/div/div/div/div/div[2]/div/div[6]/div/div/div[2]/div/div/div/div/div/div[3]/div[2]/div[1]/span").text

                # Get Discounted Price
                discounted = "None"

                # Get Alternative Part Number
                alternative = "None"

                partInfo = ("{0}, {1}, {2}, {3}".format(part, alternative, price, discounted))

            except NoSuchElementException:
                partInfo = "{0}, None, None, None".format(part)

            # Send To csv
            writeLine(self, partInfo)

    except TimeoutException or NoSuchElementException:
        writeLine(self, "Unable,to,access,website")

def stokesAP(self,driver):
    writeLine(self, "Stokes Appliance Parts")
    try:
        # Data to scrape from website
        for part in parts:
            # User Aid
            length = len(parts)
            position = (parts.index(part)) + 1
            self.lblStatus.setText("Scraping Stokes Appliance Parts {0} of {1}".format(position, length))
            self.lblStatus.setStyleSheet("color: #FF6E00")
            QtWidgets.QApplication.processEvents()

            driver.get("https://www.stokesap.com.au/")
            driver.implicitly_wait(3)

            # Insert Part Number
            driver.find_element(By.XPATH, "//*[@id='searchfilter']").send_keys(part)

            # Search
            driver.find_element(By.XPATH, "//*[@id='searchfilter']").send_keys(u'\ue007')


            # Try Except!
            try:
                # Press View
                driver.find_element(By.XPATH, "// *[ @ id = 'MoreInfo1']").send_keys(u'\ue007')

                # Get Price
                price = driver.find_element(By.XPATH, "/html/body/div[1]/div[2]/div[2]/strong/strong/div[1]/div[3]/form/div[1]/div/p").text
                price = price.replace("(inc GST)","")

                # Get Discounted Price
                discounted = "None"

                # Get Alternative Part Number
                alternative = "None"

                partInfo = ("{0}, {1}, {2}, {3}".format(part, alternative, price, discounted))

            except NoSuchElementException:
                partInfo = "{0}, None, None, None".format(part)

            # Send To csv
            writeLine(self, partInfo)

    except TimeoutException or NoSuchElementException:
        writeLine(self, "Unable,to,access,website")

def discountAP(self,driver):
    writeLine(self, "Discount Appliance Parts")
    try:
        # Data to scrape from website
        for part in parts:
            # User Aid
            length = len(parts)
            position = (parts.index(part)) + 1
            self.lblStatus.setText("Scraping Discount Appliance Parts {0} of {1}".format(position, length))
            self.lblStatus.setStyleSheet("color: #FF6E00")
            QtWidgets.QApplication.processEvents()

            driver.get("https://www.discountapplianceparts.com.au/")
            driver.implicitly_wait(3)

            # Insert Part Number
            driver.find_element(By.XPATH, "//*[@id='discount-appliance-parts-spares-online-store-australia']/header/div/div/div[2]/form/input[2]").send_keys(part)

            # Search
            driver.find_element(By.XPATH, "//*[@id='discount-appliance-parts-spares-online-store-australia']/header/div/div/div[2]/form/input[2]").send_keys(u'\ue007')

            # Try Except!
            try:
                # Get Price
                price = driver.find_element(By.XPATH, "/html/body/main/div/div/div/div/a/div[2]/span/small").text

                # Get Discounted Price
                discounted = "None"

                # Get Alternative Part Number
                alternative = "None"

                partInfo = ("{0}, {1}, {2}, {3}".format(part, alternative, price, discounted))

            except NoSuchElementException:
                partInfo = "{0}, None, None, None".format(part)

            # Send To csv
            writeLine(self, partInfo)

    except TimeoutException or NoSuchElementException:
        writeLine(self, "Unable,to,access,website")

def ovenAP(self,driver):
    writeLine(self, "Oven Parts Australia")
    try:
        # Data to scrape from website
        for part in parts:
            # User Aid
            length = len(parts)
            position = (parts.index(part)) + 1
            self.lblStatus.setText("Scraping Oven Parts Australia {0} of {1}".format(position, length))
            self.lblStatus.setStyleSheet("color: #FF6E00")
            QtWidgets.QApplication.processEvents()

            driver.get("https://www.ovenpartsaustralia.com.au/")
            driver.implicitly_wait(3)

            # Insert Part Number
            driver.find_element(By.XPATH, "//*[@id='search']").send_keys(part)

            # Search
            driver.find_element(By.XPATH, "//*[@id='search']").send_keys(u'\ue007')

            # Try Except!
            try:
                # Get Price
                price = driver.find_element(By.XPATH, "/html/body/div[4]/div/div/div/div/div[4]/div/div[2]/div[2]/div[2]/ul/li/div[1]/span/span").text

                # Get Discounted Price
                discounted = "None"

                # Get Alternative Part Number
                alternative = "None"

                partInfo = ("{0}, {1}, {2}, {3}".format(part, alternative, price, discounted))

            except NoSuchElementException:
                partInfo = "{0}, None, None, None".format(part)

            # Send To csv
            writeLine(self, partInfo)

    except TimeoutException or NoSuchElementException:
        writeLine(self, "Unable,to,access,website")

def stoveC(self,driver):
    writeLine(self, "Stove Connection")
    try:
        # Data to scrape from website
        for part in parts:
            # User Aid
            length = len(parts)
            position = (parts.index(part)) + 1
            self.lblStatus.setText("Scraping Stove Connection {0} of {1}".format(position, length))
            self.lblStatus.setStyleSheet("color: #FF6E00")
            QtWidgets.QApplication.processEvents()

            driver.get("https://stoveconnection.com.au/shop/")
            driver.implicitly_wait(3)

            # Insert Part Number
            driver.find_element(By.XPATH, "//*[@id='search']").send_keys(part)

            # Search
            driver.find_element(By.XPATH, "//*[@id='search']").send_keys(u'\ue007')

            # Try Except!
            try:
                # Get Price
                price = driver.find_element(By.XPATH, "/html/body/div[2]/main/div[3]/div[1]/div[2]/div[2]/ol/li/div/div[2]/div[1]/span/span/span").text

                # Get Discounted Price
                discounted = "None"

                # Get Alternative Part Number
                alternative = "None"

                partInfo = ("{0}, {1}, {2}, {3}".format(part, alternative, price, discounted))

            except NoSuchElementException:
                partInfo = "{0}, None, None, None".format(part)

            # Send To csv
            writeLine(self, partInfo)

    except TimeoutException or NoSuchElementException:
        writeLine(self, "Unable,to,access,website")

def ellis(self,driver):
    writeLine(self, "Ellis Electrical")
    try:
        # Data to scrape from website
        for part in parts:
            # User Aid
            length = len(parts)
            position = (parts.index(part)) + 1
            self.lblStatus.setText("Scraping Ellis Electrical {0} of {1}".format(position, length))
            self.lblStatus.setStyleSheet("color: #FF6E00")
            QtWidgets.QApplication.processEvents()

            driver.get("https://elliselectricals.com.au/")
            driver.implicitly_wait(3)

            # Insert Part Number
            driver.find_element(By.XPATH, "/html/body/div[1]/div/section[1]/div/div/div[4]/div/div/div/div/div/form/div[1]/input[1]").send_keys(part)

            # Search
            driver.find_element(By.XPATH, "/html/body/div[1]/div/section[1]/div/div/div[4]/div/div/div/div/div/form/div[1]/input[1]").send_keys(u'\ue007')

            # Try Except!
            try:
                # Get Price
                price = driver.find_element(By.XPATH, "/html/body/div[3]/div/section[1]/div/div/div[2]/div/div/div[3]/div/p/del/span/bdi").text

                # Get Discounted Price
                discounted = driver.find_element(By.XPATH, "/html/body/div[3]/div/section[1]/div/div/div[2]/div/div/div[3]/div/p/ins/span/bdi").text

                # Get Alternative Part Number
                alternative = "None"

                partInfo = ("{0}, {1}, {2}, {3}".format(part, alternative, price, discounted))

            except NoSuchElementException:
                try:
                    # Get Price
                    price = driver.find_element(By.XPATH, "/html/body/div[3]/div/section[1]/div/div/div[2]/div/div/div[3]/div/p/span/bdi").text

                    # Get Alternative Part Number
                    alternative = "None"

                    # Get Discounted Price
                    discounted = "None"

                    partInfo = "{0}, {1}, {2}, {3}".format(part, alternative, price, discounted)

                except NoSuchElementException:
                    partInfo = "{0}, None, None, None".format(part)

            # Send To csv
            writeLine(self, partInfo)

    except TimeoutException or NoSuchElementException:
        writeLine(self, "Unable,to,access,website")

def sydneyAS(self,driver):
    writeLine(self, "Sydney Appliance Service")
    try:
        # Data to scrape from website
        for part in parts:
            # User Aid
            length = len(parts)
            position = (parts.index(part)) + 1
            self.lblStatus.setText("Scraping Sydney Appliance Service {0} of {1}".format(position, length))
            self.lblStatus.setStyleSheet("color: #FF6E00")
            QtWidgets.QApplication.processEvents()

            driver.get("https://www.sydneyappliance.com.au/")
            driver.implicitly_wait(3)

            # Insert Part Number
            driver.find_element(By.XPATH, "//*[@id='RemoteSearchSuggest2']").send_keys(part)

            # Search
            driver.find_element(By.XPATH, "//*[@id='RemoteSearchSuggest2']").send_keys(u'\ue007')

            # Try Except!
            try:
                # Get Price
                price = driver.find_element(By.XPATH, "/html/body/div[1]/div[3]/div/div/div/div[3]/div[5]/table/tbody/tr/td[1]/div/span[1]").text
                price = price.replace("*","")

                # Get Discounted Price
                discounted = "None"

                # Get Alternative Part Number
                alternative = "None"

                partInfo = ("{0}, {1}, {2}, {3}".format(part, alternative, price, discounted))

            except NoSuchElementException:
                partInfo = "{0}, None, None, None".format(part)

            # Send To csv
            writeLine(self, partInfo)

    except TimeoutException or NoSuchElementException:
        writeLine(self, "Unable,to,access,website")

# Completion functions
def writeLine(self, string):
    dataFile = open("tempData.csv", "a")
    dataFile.write(string + "\n")
    dataFile.close()

    path = self.lblOutput.text()
    path = ("{0}/Scraped_data.xlsx".format(path))

    inputData = pd.read_csv("tempData.csv")
    inputData.to_excel(path, index=False)

def completeStatus(self):
    self.lblStatus.setText("Process Complete!")
    self.lblStatus.setStyleSheet("color: green")

    # Enable User Input
    for item in self.groupSites.findChildren(QtWidgets.QCheckBox):
        item.setEnabled(True)
    self.btnListBrowse.setEnabled(True)
    self.btnOutputBrowse.setEnabled(True)
    self.btnCheck.setEnabled(True)
    self.btnUncheck.setEnabled(True)
    self.btnProcess.setEnabled(True)
    self.checkProcess.setEnabled(True)
    self.input.setEnabled(True)

def fixErrorMessage():
    msg = QMessageBox()
    msg.setIcon(QMessageBox.Information)

    msg.setText("Please enter valid data where indicated red")
    msg.setWindowTitle("Error")
    msg.setStandardButtons(QMessageBox.Ok)
    retval = msg.exec_()

if __name__ == "__main__":
    import sys
    app = QtWidgets.QApplication(sys.argv)
    Form = QtWidgets.QWidget()
    ui = Ui_Form()
    ui.setupUi(Form)
    Form.show()
    sys.exit(app.exec_())

# Todo Data Presentation alterations
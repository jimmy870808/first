# -*- coding: utf-8 -*-
"""
Created on Wed Jul 22 15:46:07 2020

@author: 李俊毅
"""
# Author: Li Chun Yi
# Created Date: 2025.3.25
# Last Modified: 2025.3.25
# 
# ===================[History]==================
#   TVD_Parsing_Tool_V1.0.py
# import re

# =================[Import_Module]==============
from PyQt5.QtWidgets import (QApplication, QButtonGroup, QMessageBox, QMainWindow)
from PyQt5.QtCore import pyqtSlot

import os.path
from os.path import isfile, join
from os import listdir

import sys
import math
import datetime
import time
import pandas as pd
# import requests
# from bs4 import BeautifulSoup

from PythonExcel_V4 import PythonExcel
from ui_Designer_v2_Main import *   # Main window
from ui_Designer_v2_INI import *    # INI sub window
from ui_Designer_v2_HIL import *    # HIL sub window

# =================[Parameters]==============
"""
Bible list : Biggest number from \\10.3.0.101\\Public\\Software\\14_AutoTest\\TotalTestCaseVersion
"""

# Fold all code :  Ctrl + shift + "-"
# Expand all code :  Ctrl + shift + "+"


class MainWindow(QMainWindow, Ui_MainWindow):
    # Functions after starting
    def __init__(self, parent=None):
        super(MainWindow, self).__init__(parent)

        """ # Debugging variables """
        self.bible_list_on = True           # True : Fill in Bible list items.
        self.bible_list_on = False
        self.pivot_on = True                # True : Build pivot table in sorting page
        # self.pivot_on = False
        self.hexini_change_on = True        # True : Get hex/ini change content
        self.hexini_change_on = False
        self.sorting_page_on = True         # True : Arrange data in sorting page
        self.sorting_page_on = False
        self.screenshot_on = True           # True : Sorting page do screenshot
        self.screenshot_on = False
        self.highlight_on = True            # True : Highlight abnormal readings in Excel report
        self.highlight_on = False
        self.excel_build_on = True          # True : Open Excel report at the end
        # self.excel_build_on = False

        doe_on = True                 # True : Give default value for "fw_ver" and "type"
        # doe_on = False
        if doe_on:
            # self.fw_ver = r"TQS2QRDTXR_2.9500"
            # self.fw_ver = r"TQSTRYDTXR_2.9500"      # Fast #
            # self.fw_ver = r"TQS2QRTAXE_3.0001"
            # self.fw_ver = r"TQS2QRXXXR_3.1303"
            # self.fw_ver = r"TQS2QRXXXR_3.13"
            # self.fw_ver = "TQS2QRYUXR_2.7306"
            # self.fw_ver = "TRY2QRYUXR_2.7306"     # Local
            self.fw_ver = r"1. PreRelease\HIOL-TQS-3-QR-RE-XXXX-01.11"      # Tool

            # noinspection SpellCheckingInspection
            # self.type = "QRUA"
            self.type = "HIOL"          # Tool
            # noinspection SpellCheckingInspection
            self.hexini_path = "http://203.74.156.241:1314/redmine/projects/release-note/wiki/" \
                               "Request_2179_TQSDMRXXXE_539;" \
                               "http://203.74.156.241:1314/redmine/projects/release-note/wiki/" \
                               "Request_2615_TQSDQRXXXE_562"

        else:
            self.fw_ver = ""
            self.type = ""
            self.hexini_path = ""
            # pd.set_option('display.max_columns', None)  # Switch to print all dataframe columns

        """ # Turn these on in official mode """
        # self.bible_list_on = True  # True : Scan Bible list and fill in items.
        # self.excel_build_on = True  # True : Generate final Excel report
        # self.highlight_on = True  # True : Highlight abnormal readings in Excel report
        # self.pivot_on = True         # True : Run pivot table analysis in first page
        # self.sorting_page_on = True
        # self.fw_ver = ""
        # self.type = ""

        """ # Version related variables """
        self.TVD_manager = "ShengHung"
        # self._bible = r"D:\Coding\TVD\Ref\TotalTestCaseVersion"   # Local trial run
        self._bible = r"\\10.3.0.101\Public\Software\14_AutoTest\TotalTestCaseVersion"
        self.template = "TVD_template_V3.0.xlsx"
        self.template_path = r"\\10.3.0.101\Public\Software\14_AutoTest\2_Tool\18_TVD_parsing_tool\3.Template"
        # Valid path on NAS
        # noinspection SpellCheckingInspection
        self.path_list = [
            r"D:\Coding\TVD\Data_path\S29_VCOL\2. Release",
            r'\\10.3.0.101\Public\Software\3_FW_release_record\Hex\TQS\MB\1.TMR',
            r'\\10.3.0.101\Public\Software\3_FW_release_record\Hex\TQS\MB\1.MOH',
            r'\\10.3.0.101\Public\Software\3_FW_release_record\Hex\TQS\MB\1.QRUA\2.TQS3\1.TQS3',
            r'\\10.3.0.101\Public\Software\3_FW_release_record\Hex\TQS\MB\1.QRUA\2.TQS3\3.TQSK',
            r'\\10.3.0.101\Public\Software\3_FW_release_record\Hex\TQS\MB\2.QL-Plastic',
            r'\\10.3.0.101\Public\Software\3_FW_release_record\Hex\TQS\MB\1.QRUA',
            r'\\10.3.0.101\Public\Software\3_FW_release_record\Hex\HCS2\App',
            r'\\10.3.0.101\Public\Software\3_FW_release_record\Hex\HCS\App',
            r'\\10.3.0.50\Public\Software\3_FW_release_record\Hex\TQS\MB\1.QRUA\5.TQSC\1.QRUA']
        self.tool_list = [
            ["EMFT", r"\\10.3.0.101\Public\Software\4_Tool\3.Support_產線工具\S11_Homemade_EMFT"],
            ["UpgradeTool", r"\\10.3.0.101\Public\Software\4_Tool\3.Support_產線工具\S17_Homemade_UpgradeTool"],
            ["CAN2UART", r"\\10.3.0.101\public\Software\3_FW_release_record\Hex\TQS\CAN2UART\CAN2UART_APP"],
            ["PCBA", r"\\10.3.0.101\Public\Software\3_FW_release_record\Hex\TQS\MB\5.PCBA_TestCode"],
            ["PCBA_HCS", r"\\10.3.0.101\Public\Software\3_FW_release_record\Hex\HCS\TestCode"],
            ["ET", r"\\10.3.0.101\Public\Software\4_Tool\3.Support_產線工具\S12_Homemade_ETOL"],
            ["ICBF", r"\\10.3.0.101\Public\Software\4_Tool\3.Support_產線工具\S07_Outsourcing_Programming_ICBF"],
            ["HIOL", r"\\10.3.0.101\Public\Software\4_Tool\3.Support_產線工具\S26_HIOL"],
            ["DSTool", r"\\10.3.0.101\Public\Software\4_Tool\3.Support_產線工具\S27_Homemade_DSTool"],
            ["OnSiteProg", r"\\10.3.0.101\Public\Software\4_Tool\3.Support_產線工具\S25_OnSiteProg\2.Release"],
            ["VCOL", r"\\10.3.0.101\Public\Software\4_Tool\3.Support_產線工具\S29_VCOL"],
            ["DCOL", r"\\10.3.0.101\Public\Software\4_Tool\3.Support_產線工具\S30_DCOL"],
            ["OSCT", r"\\10.3.0.101\Public\Software\4_Tool\3.Support_產線工具\S33_OSCT"],
            ["CAOL", r"\\10.3.0.101\Public\Software\4_Tool\3.Support_產線工具\S21_Homemade_UDSUpgradeTool"]
        ]

        """ # Fixed variables """
        # Parsing step variables
        self.set_owner = None               # Sorting page : Count owner
        self.set_reviewer = None            # Sorting page : Count reviewer
        self.tool_type_ = None              # Tool case : Get special type for bible list searching
        self.ini_window = True              # HIL+INI case : False, INI case : True
        self.sub_folder_list = None         # HIL+INI : Feed list from HIL to INI
        self.sub_path_list = None           # HIL+INI : Feed list from HIL to INI
        self.Final_release_fw_ver = None    # Sorting page : FW ver
        self.Final_release_ini_ver = None   # Sorting page : INI ver
        self.RD_name = None                 # Sorting page : RD name
        self.git_tag = None                 # Sorting page : Git tag for screenshot
        self.summary = [0, 0, 0, 0]         # Build excel : List to record PASS/FAIL/WARNING/NA
        self.bible_output = []              # Bible list : FW/INI counts from empty, HIL counts from HIL output list
        self.bible_list_full = None         # Bible list : Latest bible list (full path)
        self.file_name_txt = None           # Message update : Write log into txt

        # Path variables
        self.code_path = os.getcwd()        # Python code location (Avoid calling os function several time)
        # Save bible list in csv to release memory
        self.bible_list_csv = self.code_path + r"\SWtestParsingReport\bible_list.csv"
        self.tool_path = self.code_path + r"\Tool"     # Path for other tool

        # Sub window functions
        self.sub_signal = False
        self.sub_list = []
        self.sub_items = []

        # Fixed list/string/dataframe variables
        self.delay_time = 1
        # noinspection SpellCheckingInspection
        self.test_items_keyword = [         # TVD parsing report column items
            "Customer", "Redmine", "韌體", "ini", "工具", "Classification", "Task", "Test Case No", "Test Case",
            "Test Item", "Test Q'ty", "Test Result", "Switch Option", "EWM", "RWM", "Duration", "Owner", "Reviewer",
            "Project", "ReportVersion", "Department", "FilePath", "Test case of bible list"
        ]
        self.df_empty = self.empty_list(self.test_items_keyword)  # Define empty list for Excel
        self.df_empty_single = self.empty_list(["Single_title"])
        self.status_first = "Step 0 of 7 : User input ..."      # First status shown

        # UI variables (Program start)
        self.setupUi(self)  # Call function from ui_Designer_v2
        self.init_ui()

    def init_ui(self):
        """ # Default box value """
        fw_ver_ = self.fw_ver
        type_ = self.type
        status_ = self.status_first
        hex_path = self.hexini_path

        """ # Set boxes for inputting info """
        self.my_line_edit1.setText(fw_ver_)  # Default text
        self.my_line_edit2.setText(type_)  # Default text
        self.my_line_edit4.setText(status_)  # Default text
        self.my_line_edit6.setText(hex_path)  # Default text

        """ # Searching Bible list """
        self.bible_list_full = self.bible_list_get(self._bible)  # Get bible list
        self.my_line_edit5.setText(self.bible_list_full.split("\\")[-1])  # Default text

        """ # Set button for data selecting mode """
        # self.radioButton_1.setChecked(True)
        # self.radioButton_2.setChecked(False)
        QButtonGroup(self).addButton(self.radioButton_1, 1)
        QButtonGroup(self).addButton(self.radioButton_2, 2)

    def initial_all(self):
        """ # Fixed variables """
        # Parsing step variables
        self.set_owner = None  # Sorting page : Count owner
        self.set_reviewer = None  # Sorting page : Count reviewer
        self.tool_type_ = None  # Tool case : Get special type for bible list searching
        self.ini_window = True  # HIL+INI case : False, INI case : True
        self.sub_folder_list = None  # HIL+INI : Feed list from HIL to INI
        self.sub_path_list = None  # HIL+INI : Feed list from HIL to INI
        self.Final_release_fw_ver = None  # Sorting page : FW ver
        self.Final_release_ini_ver = None  # Sorting page : INI ver
        self.RD_name = None  # Sorting page : RD name
        self.git_tag = None  # Sorting page : Git tag for screenshot
        self.summary = [0, 0, 0, 0]  # Build excel : List to record PASS/FAIL/WARNING/NA
        self.bible_output = []  # Bible list : FW/INI counts from empty, HIL counts from HIL output list
        self.file_name_txt = None  # Message update : Write log into txt

        # Sub window functions
        self.sub_signal = False
        self.sub_list = []
        self.sub_items = []

    def version_check(self):  # 1. Template matched : keep running 2. Not matched : show error
        """ # Check tool version via template version """
        template_version = os.listdir(self.template_path)  # Get all file name from template folder path
        # print(template_version[0])                        # Supposed to be only one template file on server
        if template_version[0] != self.template:
            version = template_version[0].split("_")[-1].split(".xl")[0]  # TVD_template_V1.0.xlsx => V1.0
            QMessageBox.warning(self, "Tool version not matched !", "Please update to tool " + version)
            return False
        else:
            return True

    # Initializing functions
    @staticmethod
    def empty_list(input_title):  # Define empty list for Excel
        """ # Define empty dataframe with headers """
        df_empty = pd.DataFrame(columns=input_title, index=list(range(0, 1)))
        for row in range(len(df_empty)):  # Cleaning dataframe from "NAN" to ""
            for column in range(len(df_empty.columns)):
                # df_empty.iloc[row][column] = ""          # Original
                # df_empty.loc[row, column] = ""  # New
                df_empty.iloc[row, column] = ""  # New
        return df_empty

    def bible_list_get(self, path):
        # Input bible list path, output latest bible list file(full path)

        """ # Glob function investment """
        # path = r"D:\Coding\TVD\Ref\TotalTestCaseVersion"                    # Define path for bible list
        #
        # print(path)
        # print((glob.glob(os.path.join(path, "*"))))    # Print all items under folder path
        # print((glob.glob(os.path.join(path, r"*SWTestList_V"))))  # Print all items under folder path

        """ # Scan all version numbers from Bible list folder """
        self.message_update("Scanning Bible list from path", True)
        self.message_update(self._bible, False)
        name_list = [f for f in listdir(path) if isfile(join(path, f))]  # Scan all file from bible folder
        ver_list = []  # Empty list for inputting version number
        for i in range(len(name_list)):
            name = name_list[i].split(".")[0]  # SWTestList_V259.xlsx => SWTestList_V259
            version = name.split("V")[-1]  # SWTestList_V259 => 259
            # if str(version).isdigit() == True:  # Only get int value(Normal expression)
            if str(version).isdigit():  # Only get int value (Better expression)
                ver_list.append(version)  # Put value into empty list
        # print(ver_list)

        """ # Getting latest Bible list version """
        current_ver = ver_list[0]
        for i in range(len(ver_list)):
            if int(ver_list[i]) > int(current_ver):
                # print(current_ver, "<", ver_list[i])
                current_ver = ver_list[i]
        # print(current_ver)
        current_ver = path + "\\SWTestList_V" + str(current_ver) + ".xlsx"
        self.message_update("Found Bible list", True)
        self.message_update(current_ver, False)

        return current_ver  # Return "D:\Coding\TVD\Ref\TotalTestCaseVersion\SWTestList_V262.xlsx"

    # Common functions
    def message_update(self, message, stamp):
        """ # Get current time """
        width = 26
        if stamp:  # If stamp == True, show time stamp + Message
            now = datetime.datetime.now()  # Get current day and time
            # print(now.strftime("%Y.%m.%d %H:%M:%S"))  # 2025.03.12 16:29:24
            time_stamp = "[ " + now.strftime("%Y.%m.%d %H:%M:%S") + " ]"  # [ 2025.03.12 16:29:24 ]
            message = time_stamp.ljust(width) + message  # Message add timestamp and 26 blanks
        else:
            message = " ".ljust(width) + message  # Message add 26 blanks
        self.my_line_edit3.append(message + "\n")  # Write message into UI box via "append"
        QApplication.processEvents()  # Force program to wait until the time-consuming command is finished

        """ # Show message and save into txt file """
        if len(message) > 0:
            print(message)
            if str(self.file_name_txt) != "None":
                with open(self.file_name_txt, 'a') as f:
                    f.write(message + '\n\n')
                    f.close()

    def status_update(self, message):
        max_str = 110
        if len(message) > max_str:
            message = message[(-1 - max_str):-1]

        """ # Show message and save into txt file """
        self.my_line_edit4.setText(message)  # Replace message into status UI box via "setText"
        QApplication.processEvents()  # Force program to wait until the time-consuming command is finished

    # Button functions
    @pyqtSlot()
    def on_button_fw_clicked(self):
        # Actions after clicking "Parsing FW report"
        """ # Record user action """
        self.message_update("<< User clicked : Parsing FW Report >>", True)

        """ # Get box value """
        fw_ver_, type_, hexini_, input_valid = self.get_box_value()
        if not input_valid:  # Stop if box value is invalid
            return

        """ # Remind user before killing Excel """
        if not self.excel_kill():
            return

        """ # Disable all buttons while parsing report """
        self.button_fw.setDisabled(True)
        self.button_ini.setDisabled(True)
        self.button_hil.setDisabled(True)
        self.button_tool.setDisabled(True)

        """ # Define txt and xlsx report name by timestamp """
        file_name_excel, file_name_excel_abs = self.timestamp_name(fw_ver_, "FW")

        """ # Create folder and build excel """
        self.excel_build(file_name_excel)

        """ # Feedback fw_ver, type, bible ver, sub folder input(INI case) """
        self.message_update("Parsing FW report", True)
        self.message_update("FW Ver     : " + fw_ver_, False)
        self.message_update("Type       : " + type_, False)
        self.message_update("Bible Ver  : " + self.bible_list_full.split("\\")[-1], False)

        """ # Get xlsx path from fw_ver """
        path_ = self.path_searching(fw_ver_, "2.TestData\\")  # Get folder path via folder name
        if len(path_) > 0:
            # Get GitTag for screenshot(If TQS2QRXXXR_3.13, get TQS2QRXXXR_3.1303)
            self.git_tag = path_[-1].split("\\")[-3]
            # D:\Coding\TVD\Data_path\S29_VCOL\2. Release\TQSTRYDTXR_2.9601\2.TestData\ => TQSTRYDTXR_2.9601

            for i in path_:
                self.message_update(i, False)
        else:
            """ # Enable all buttons """
            self.button_fw.setDisabled(False)
            self.button_ini.setDisabled(False)
            self.button_hil.setDisabled(False)
            self.button_tool.setDisabled(False)
            return

        """ # FW starts from 2.1 to 2.5 """
        # folder_list = ["2.1.Tessy", "2.2.PolySpace", "2.3.InternalTest", "2.4.SimulatorTest", "2.5.SystemTestReport"]
        folder_list = ["2.1.Tessy", "2.2.PolySpace", "2.3.InternalTest", "2.4.SimulatorTest",
                       "2.5.SystemTestReport", "2.6.iniTestReport"]

        if len(self.bible_output) == 0:     # Only FW case
            bible_output = self.folder_parsing("FW", folder_list, path_, type_, file_name_excel_abs, "")
        else:                               # HIL + FW case
            temp_ = self.folder_parsing("FW", folder_list, path_, type_, file_name_excel_abs, "")
            # for i in range(len(temp_)):
            #     print(temp_[i])
            #     self.bible_output.append(temp_[i])
            self.bible_output = temp_
            bible_output = self.bible_output

        """ # Count types and build parameter column """
        self.para_col_build(bible_output, fw_ver_, type_, hexini_, file_name_excel_abs)

        """ # Build Bible list items """
        if self.bible_list_on:
            # list_duration = self.bible_list_build(bible_output, type_, file_name_excel_abs)
            self.bible_list_build(bible_output, type_, file_name_excel_abs)

        """ # Highlight abnormal results """
        if self.highlight_on:
            self.highlight(file_name_excel_abs, len(bible_output) + 1, 25)

        """ # Pivot table analysis """
        if self.pivot_on:
            self.pivot_analysis(file_name_excel_abs, len(bible_output))

        """ # Hex/Ini change checking """
        if self.hexini_change_on:
            self.hexini_change(file_name_excel_abs, hexini_)

        """ # Sorting page arrangement """
        if self.sorting_page_on:
            self.sorting_page_arrange(file_name_excel_abs)

        """ # Screen shot """
        if self.screenshot_on:
            self.screenshot_run(file_name_excel_abs)

        """ # Show Excel report once finished """
        if self.excel_build_on:
            """ # Open excel """
            __excel_handle = PythonExcel(file_name_excel_abs)
            __excel_handle.open_excel(file_name_excel_abs, 600, 500)  # Open excel, size = 600 x 500
            # del __excel_handle
            time.sleep(self.delay_time)
            """ # Feedback xlsx report status """
            self.message_update("Excel file successfully generated :", True)
            self.status_update("Excel file successfully generated")
            self.message_update("\\" + file_name_excel, False)
            self.message_update("Pass Count(Pass + Warning) : " + str(self.summary[0] + self.summary[2]), False)
            self.message_update("Fail Count : " + str(self.summary[1]), False)
            self.message_update("Warning : " + str(self.summary[2]), False)
            self.message_update("NA : " + str(self.summary[3]), False)
            self.summary = [0, 0, 0, 0]  # Initialize test result recording list(PASS FAIL WARNING NA)
            QMessageBox.information(self, "Info", "Excel file successfully generated!!!\n\n"
                                    + self.code_path + "\n" + "\\" + file_name_excel)

        """ # Enable all buttons """
        self.button_fw.setDisabled(False)
        self.button_ini.setDisabled(False)
        self.button_hil.setDisabled(False)
        self.button_tool.setDisabled(False)

        """ # Initialize all """
        self.initial_all()

    @pyqtSlot()
    def on_button_ini_clicked(self):
        # Actions after clicking "Parsing INI report"
        """ # Record user action """
        self.message_update("<< User clicked : Parsing INI Report >>", True)

        """ # Get box value """
        fw_ver_, type_, hexini_, input_valid = self.get_box_value()
        if not input_valid:  # Stop if box value is invalid
            return

        """ # If only INI : run sub window. If HIL+INI, dont run sub window (Already run in HIL) """
        if self.ini_window:
            """ # Remind user before killing Excel """
            if not self.excel_kill():
                return

            """ # Disable all buttons while parsing report """
            self.button_fw.setDisabled(True)
            self.button_ini.setDisabled(True)
            self.button_hil.setDisabled(True)
            self.button_tool.setDisabled(True)

            """ # Sub window : Run """
            dialog = SubWindowINI()
            # Integrate with signal emit function
            dialog.dialog_signal.connect(self.signal_receiver)
            dialog.show()
            dialog.exec_()

            """ # Sub window : Return if no """
            if not self.sub_signal:
                # Enable all buttons
                self.button_fw.setDisabled(False)
                self.button_ini.setDisabled(False)
                self.button_hil.setDisabled(False)
                self.button_tool.setDisabled(False)
                return

            """ # Sub window : Check invalid path """
            sub_path_list = []  # Recording ini input path name
            sub_folder_list = []  # Recording ini input folder name
            print("INI parsing items : ")
            for i in range(len(self.sub_list)):  # Check all input from sub window
                print(self.sub_list[i])
                if len(self.sub_list[i].split("\\")) == 1:  # If item does not contain "\", record to folder list
                    sub_folder_list.append(self.sub_list[i])
                else:  # If contains "\" and existing, record to path list
                    if not os.path.isfile(self.sub_list[i]):  # Warning if path not exist
                        QMessageBox.warning(self, "Warning ! Path not existing ! ", '"' + str(self.sub_list[i][1:-1]) + '"')
                        self.message_update("Warning ! Path not existing ! ", False)
                        self.message_update(str(self.sub_list[i][1:-1]), False)
                    else:
                        sub_path_list.append(self.sub_list[i])
        else:
            sub_folder_list = self.sub_folder_list
            sub_path_list = self.sub_path_list

        """ # Define txt and xlsx report name by timestamp """
        file_name_excel, file_name_excel_abs = self.timestamp_name(self.sub_items[0], "INI")

        """ # Create folder and build excel """
        self.excel_build(file_name_excel)

        """ # Feedback fw_ver, type, bible ver, sub folder input(INI case) """
        self.message_update("Parsing INI report", True)
        self.message_update("FW Ver     : " + fw_ver_, False)
        self.message_update("Type       : " + type_, False)
        self.message_update("Bible Ver  : " + self.bible_list_full.split("\\")[-1], False)
        self.message_update("Ini input (all)     : ", False)
        for i in range(len(self.sub_list)):
            self.message_update(self.sub_list[i], False)
        self.message_update("Ini input (folder)     : ", False)
        for i in range(len(sub_folder_list)):
            self.message_update(sub_folder_list[i], False)
        self.message_update("Ini input (path)       : ", False)
        for i in range(len(sub_path_list)):
            self.message_update(sub_path_list[i][1:-1], False)  # [1:-1] : Delete first alphabet to avoid splitting

        """ # Get xlsx path from fw_ver """
        path_ = self.path_searching(fw_ver_, "2.TestData\\")  # Get folder path via folder name
        if len(path_) > 0:
            # Get GitTag for screenshot(If TQS2QRXXXR_3.13, get TQS2QRXXXR_3.1303)
            self.git_tag = path_[-1].split("\\")[-3]
            # D:\Coding\TVD\Data_path\S29_VCOL\2. Release\TQSTRYDTXR_2.9601\2.TestData\ => TQSTRYDTXR_2.9601
            for i in path_:
                self.message_update(i, False)
        else:
            """ # Enable all buttons """
            self.button_fw.setDisabled(False)
            self.button_ini.setDisabled(False)
            self.button_hil.setDisabled(False)
            self.button_tool.setDisabled(False)
            return

        """ # INI runs 2.6 """
        folder_list = sub_folder_list
        # folder_list = ["2.6.iniTestReport"]

        if len(self.bible_output) == 0:     # Only INI case
            bible_output = self.folder_parsing("INI", folder_list, path_, type_, file_name_excel_abs, sub_path_list)
        else:                               # HIL + INI case
            temp_ = self.folder_parsing("INI", folder_list, path_, type_, file_name_excel_abs, sub_path_list)
            # for i in range(len(temp_)):
            #     print(temp_[i])
            #     self.bible_output.append(temp_[i])
            self.bible_output = temp_
            bible_output = self.bible_output

        """ # Count types and build parameter column """
        print(fw_ver_, type_, file_name_excel_abs)
        self.para_col_build(bible_output, fw_ver_, type_, hexini_, file_name_excel_abs)

        """ # Build Bible list items """
        # for i in range(len(bible_output)):
        #     print(bible_output[i])
        if self.bible_list_on:
            # list_duration = self.bible_list_build(bible_output, type_, file_name_excel_abs)
            self.bible_list_build(bible_output, type_, file_name_excel_abs)

        """ # Highlight abnormal results """
        if self.highlight_on:
            self.highlight(file_name_excel_abs, len(bible_output) + 1, 25)

        """ # Pivot table analysis """
        if self.pivot_on:
            self.pivot_analysis(file_name_excel_abs, len(bible_output))

        """ # Hex/Ini change checking """
        if self.hexini_change_on:
            self.hexini_change(file_name_excel_abs, hexini_)

        """ # Sorting page arrangement """
        if self.sorting_page_on:
            self.sorting_page_arrange(file_name_excel_abs)

        """ # Screen shot """
        if self.screenshot_on:
            self.screenshot_run(file_name_excel_abs)

        """ # Show Excel report once finished """
        if self.excel_build_on:
            """ # Open excel """
            __excel_handle = PythonExcel(file_name_excel_abs)
            # Adjust first column width
            __excel_handle.column_width("Sheet1", 1, 30)
            __excel_handle.open_excel(file_name_excel_abs, 600, 500)  # Open excel, size = 600 x 500
            # del __excel_handle
            time.sleep(self.delay_time)
            """ # Feedback xlsx report status """
            self.message_update("Excel file successfully generated :", True)
            self.status_update("Step 7 of 7 : Excel file successfully generated")
            self.message_update("\\" + file_name_excel, False)
            self.message_update("Pass Count(Pass + Warning) : " + str(self.summary[0] + self.summary[2]), False)
            self.message_update("Fail Count : " + str(self.summary[1]), False)
            self.message_update("Warning : " + str(self.summary[2]), False)
            self.message_update("NA : " + str(self.summary[3]), False)
            self.summary = [0, 0, 0, 0]  # Initialize test result recording list(PASS FAIL WARNING NA)
            QMessageBox.information(self, "Info", "Excel file successfully generated!!!\n\n"
                                    + self.code_path + "\n" + "\\" + file_name_excel)

        # print(" ")
        # print("<< Pressed >> Parsing INI Report")
        # if QButtonGroup(self).checkedId() == 1:
        #     self.message_update("Mode : On-line", True)
        # elif QButtonGroup(self).checkedId() == 2:
        #     self.message_update("Mode : Off-line", True)

        """ # Initializing signals from sub window """
        self.sub_signal = False
        self.sub_list = []
        self.sub_items = []

        """ # Enable all buttons """
        self.button_fw.setDisabled(False)
        self.button_ini.setDisabled(False)
        self.button_hil.setDisabled(False)
        self.button_tool.setDisabled(False)

        """ # Initialize all """
        self.initial_all()

    @pyqtSlot()
    def on_button_tool_clicked(self):
        # Actions after clicking "Parsing Tool report"
        """ # Record user action """
        self.message_update("<< User clicked : Parsing Tool Report >>", True)

        """ # Get box value """
        fw_ver_, type_, hexini_, input_valid = self.get_box_value()
        if not input_valid:  # Stop if box value is invalid
            return

        """ # Remind user before killing Excel """
        if not self.excel_kill():
            return

        """ # Get xlsx path from fw_ver & type, check if path is valid """
        type_match = False
        # If input type : "EMFT", FW_ver = "1. PreRelease_TQS2\EMFT-XXX-2-XX-RE-XXXX-53.51"
        # Output path : "\\10.3.0.101\Public\Software\4_Tool\3.Support_產線工具\S11_Homemade_EMFT
        # \1. PreRelease_TQS2\EMFT-XXX-2-XX-RE-XXXX-53.51\2.TestData\"
        for i in range(len(self.tool_list)):
            list_type = self.tool_list[i][0]
            list_path = self.tool_list[i][1]

            if type_ == list_type:
                path_ = [list_path + "\\" + fw_ver_ + "\\" + "2.TestData" + "\\"]
                type_match = True
                if type_ == "ET":       # VBA "type_ET" case(HCS, EWP, CLS)
                    self.tool_type_ = fw_ver_.split("\\")[0].split("-")[1]
                    # Input "ETOL-CLS-4\1. PreRelease\V00.03", output "CLS"
                else:
                    self.tool_type_ = "TQS"
                break

        # path_ = [r"\\10.3.0.101\Public\Software\4_Tool\3.Support_產線工具\S26_HIOL\1. PreRelease
        # \HIOL-TQS-2-QR-RE-XXXX-00.01\2.TestData\\"]
        if type_match:
            self.message_update("Tool report running path : " + path_[0], True)
        else:
            self.message_update("Invalid input type : " + type_, True)
            QMessageBox.warning(self, "Warning", 'Wrong "Type" for tool report, please check again.')
            return

        """ # Disable all buttons while parsing report """
        self.button_fw.setDisabled(True)
        self.button_ini.setDisabled(True)
        self.button_hil.setDisabled(True)
        self.button_tool.setDisabled(True)

        """ # Define txt and xlsx report name by timestamp """
        file_name_excel, file_name_excel_abs = self.timestamp_name(fw_ver_, "TOOL")

        """ # Create folder and build excel """
        self.excel_build(file_name_excel)

        """ # Feedback fw_ver, type, bible ver, sub folder input(INI case) """
        self.message_update("Parsing Tool report", True)
        self.message_update("FW Ver     : " + fw_ver_, False)
        self.message_update("Type       : " + type_, False)
        self.message_update("Bible Ver  : " + self.bible_list_full.split("\\")[-1], False)

        """ # Tool no need to scan NAS folder """
        # folder_list = ["2.1.Tessy", "2.2.PolySpace", "2.3.InternalTest", "2.4.SimulatorTest", "2.5.SystemTestReport"]
        folder_list = [""]
        bible_output = self.folder_parsing("TOOL", folder_list, path_, type_, file_name_excel_abs, "")


        """ # Count types and build parameter column """
        self.para_col_build(bible_output, fw_ver_, type_, hexini_, file_name_excel_abs)

        """ # Build Bible list items """
        if self.bible_list_on:
            # list_duration = self.bible_list_build(bible_output, type_, file_name_excel_abs)
            self.bible_list_build(bible_output, type_, file_name_excel_abs)

        """ # Highlight abnormal results """
        if self.highlight_on:
            self.highlight(file_name_excel_abs, len(bible_output) + 1, 25)

        """ # Pivot table analysis """
        if self.pivot_on:
            self.pivot_analysis(file_name_excel_abs, len(bible_output))

        """ # Show Excel report once finished """
        if self.excel_build_on:
            """ # Open excel """
            __excel_handle = PythonExcel(file_name_excel_abs)
            __excel_handle.open_excel(file_name_excel_abs, 600, 500)  # Open excel, size = 600 x 500
            # del __excel_handle
            time.sleep(self.delay_time)
            """ # Feedback xlsx report status """
            self.message_update("Excel file successfully generated :", True)
            self.status_update("Excel file successfully generated")
            self.message_update("\\" + file_name_excel, False)
            self.message_update("Pass Count(Pass + Warning) : " + str(self.summary[0] + self.summary[2]), False)
            self.message_update("Fail Count : " + str(self.summary[1]), False)
            self.message_update("Warning : " + str(self.summary[2]), False)
            self.message_update("NA : " + str(self.summary[3]), False)
            self.summary = [0, 0, 0, 0]  # Initialize test result recording list(PASS FAIL WARNING NA)
            QMessageBox.information(self, "Info", "Excel file successfully generated!!!\n\n"
                                    + self.code_path + "\n" + "\\" + file_name_excel)

        """ # Enable all buttons """
        self.button_fw.setDisabled(False)
        self.button_ini.setDisabled(False)
        self.button_hil.setDisabled(False)
        self.button_tool.setDisabled(False)

        """ # Initialize all """
        self.initial_all()

    @pyqtSlot()
    def on_button_hil_clicked(self):
        # Actions after clicking "Parsing HIL report"
        """ # Record user action """
        self.message_update("<< User clicked : Parsing HIL Report >>", True)

        """ # Get box value """
        fw_ver_, type_, hexini_, input_valid = self.get_box_value()
        if not input_valid:  # Stop if box value is invalid
            return

        """ # Remind user before killing Excel """
        if not self.excel_kill():
            return

        """ # Disable all buttons """
        self.button_fw.setDisabled(True)
        self.button_ini.setDisabled(True)
        self.button_hil.setDisabled(True)
        self.button_tool.setDisabled(True)

        """ # Sub window : Run """
        dialog = SubWindowHIL()
        # Integrate with signal emit function
        dialog.dialog_signal.connect(self.signal_receiver)
        dialog.show()
        dialog.exec_()

        """ # Sub window : Return if no """
        if not self.sub_signal:
            # Enable all buttons
            self.button_fw.setDisabled(False)
            self.button_ini.setDisabled(False)
            self.button_hil.setDisabled(False)
            self.button_tool.setDisabled(False)
            return
        else:
            hil_customer = self.sub_items[0]
            hil_redmine = self.sub_items[1]
            hil_reviewer = self.sub_items[2]
            report_add = self.sub_items[4]
            hil_list = self.sub_list

        """ # If user select INI : INI case """
        if report_add == "INI":
            self.ini_window = False
            """ # Get INI path from sub window """
            dialog = SubWindowINI()
            # Integrate with signal emit function
            dialog.dialog_signal.connect(self.signal_receiver)
            dialog.show()
            dialog.exec_()

            """ # Sub window : Return if no """
            if not self.sub_signal:
                # Enable all buttons
                self.button_fw.setDisabled(False)
                self.button_ini.setDisabled(False)
                self.button_hil.setDisabled(False)
                self.button_tool.setDisabled(False)
                return

            """ # Sub window : Check invalid path """
            sub_path_list = []  # Recording ini input path name
            sub_folder_list = []  # Recording ini input folder name
            print("INI parsing items : ")
            for i in range(len(self.sub_list)):  # Check all input from sub window
                print(self.sub_list[i])
                if len(self.sub_list[i].split("\\")) == 1:  # If item does not contain "\", record to folder list
                    sub_folder_list.append(self.sub_list[i])
                else:  # If contains "\" and existing, record to path list
                    if not os.path.isfile(self.sub_list[i]):  # Warning if path not exist
                        QMessageBox.warning(self, "Warning ! Path not existing ! ", '"' + str(self.sub_list[i][1:-1]) + '"')
                        self.message_update("Warning ! Path not existing ! ", False)
                        self.message_update(str(self.sub_list[i][1:-1]), False)
                    else:
                        sub_path_list.append(self.sub_list[i])
            self.sub_path_list = sub_path_list          # Feed to INI process
            self.sub_folder_list = sub_folder_list      # Feed to INI process

        """ # Define txt and xlsx report name by timestamp """
        file_name_excel = r"SWtestParsingReport\\HIL.xlsx"
        self.file_name_txt = r"SWtestParsingReport\\HIL.txt"
        file_name_excel_abs = self.code_path + "\\" + file_name_excel

        """ # Create folder and build excel """
        self.excel_build(file_name_excel)

        """ # Feedback fw_ver, type, bible ver, sub folder input(HIL case) """

        self.message_update("Parsing HIL report", True)
        self.message_update("Customer   : " + hil_customer, False)
        self.message_update("Redmine No : " + hil_redmine, False)
        self.message_update("Reviewer   : " + hil_reviewer, False)
        self.message_update("Bible Ver  : " + self.bible_list_full.split("\\")[-1], False)

        """ # Get HTML path from sub window """
        path_ = []  # Recording ini input path name
        for i in range(len(hil_list)):  # Check all input from sub window
            if len(hil_list[i].split("\\")) != 1:  # If contains "\" and existing, record to path list
                path_.append(hil_list[i])
        if len(path_) > 0:
            for i in path_:
                self.message_update(i, False)
        else:
            """ # Enable all buttons """
            self.button_fw.setDisabled(False)
            self.button_ini.setDisabled(False)
            self.button_hil.setDisabled(False)
            self.button_tool.setDisabled(False)
            return

        """ # Parsing HIL report """
        bible_output = []  # Define empty dataframe for bible list searching
        df_output = self.html_parsing(
            path_, hil_customer=hil_customer, hil_redmine=hil_redmine, hil_reviewer=hil_reviewer)

        """ # Add two empty columns for parameters """
        df_output.insert(0, 'Ref_1', "", allow_duplicates=False)  # Add empty row for parameters
        df_output.insert(1, 'Ref_2', "", allow_duplicates=False)  # Add empty row for parameters

        """ # Transfer dataframe to Excel """
        if self.excel_build_on:
            self.message_update("Generating xlsx report, please wait for a while . . . . . .", True)
            self.status_update("Step 4 of 7 : Generating xlsx report . . .")
            # Transfer dataframe to final excel report
            # self.build_xlsx(df_output, file_name_excel, file_name_excel_abs)

            # Refresh path and call win32com function again
            __excel_handle = PythonExcel(file_name_excel_abs)
            if len(bible_output) != 0:
                __excel_handle.write_pandas(df_output, "Sheet1", 2 + len(bible_output), 1, False)
                # self.build_xlsx(df_output, file_name_excel_abs, 2 + len(bible_output), 1)
            else:
                __excel_handle.write_pandas(df_output, "Sheet1", 2, 1, False)
                # self.build_xlsx(df_output, file_name_excel_abs, 2, 1)
            del __excel_handle
            time.sleep(self.delay_time)

        # Save df_output into bible_output
        for i in range(len(df_output)):
            # df_output.iloc[i, 4] : FW ver.
            # df_output.iloc[i, 5] : INI ver.
            # df_output.iloc[i, 9] : Test Case No
            # df_output.iloc[i, 16] : RWM
            array_ = [df_output.iloc[i, 4], df_output.iloc[i, 5], df_output.iloc[i, 9], df_output.iloc[i, 16]]
            bible_output.append(array_)

        self.bible_output = bible_output

        """ # FW case """
        if report_add == "FW":
            self.on_button_fw_clicked()

        """ # INI case """
        if report_add == "INI":
            self.on_button_ini_clicked()

        """ # Initialize all """
        self.bible_output = []

    def get_box_value(self):  # 1. Read (fw_ver/type) value 2. Warning if weird
        """ # Get box value """
        fw_ver_ = self.my_line_edit1.text()
        type_ = self.my_line_edit2.text()
        hexini_ = self.my_line_edit6.toPlainText()

        """ # Warning if input is weird """
        if len(str(fw_ver_)) < 11:
            QMessageBox.warning(self, "Warning", "Invalid input from FW Ver.")
            return fw_ver_, type_, hexini_, False
        if len(str(type_)) < 2:
            QMessageBox.warning(self, "Warning", "Invalid input from Type")
            return fw_ver_, type_, hexini_, False
        if len(str(hexini_)) < 1:
            QMessageBox.warning(self, "Warning", "Please input hex/ini release note path")
            return fw_ver_, type_, hexini_, False
        return fw_ver_, type_, hexini_, True

    def excel_kill(self):  # 1. Inform user to close Excel Windows 2. Kill Excel
        if not os.path.isfile(r"SWtestParsingReport\\HIL.xlsx"):  # No HIL file = FW/INI case
            # """ # Remind user before killing Excel """
            msgbox = QMessageBox(self)
            msgbox.setWindowTitle("Killing all excel")
            msgbox.setText('Already closed all Excel windows ?')
            msgbox.setStandardButtons(QMessageBox.Yes | QMessageBox.Cancel)
            msgbox.setDefaultButton(QMessageBox.Cancel)
            ret = msgbox.exec() == QMessageBox.Yes
            if ret:
                self.message_update("Killing all excel", True)
                os.system("taskkill /f /t /im EXCEL.EXE")  # Kill all Excel files in background
            return ret
        else:
            return True

            # """ # Remind user before killing Excel """
            # mbox = QtWidgets.QMessageBox(self)  # Show box
            # mbox.setText("Already close all Excel windows ?" + '\n(All Excel in background will be killed!!!)')
            # # Content
            # mbox.addButton("Yes, please run parsing tool ! ", 3)  # Selection A
            # mbox.addButton('No', 3)  # Selection B
            # ret = mbox.exec()  # 0 : Kill.  1 : Dont kill
            #
            # if ret == 0:  # Kill all Excel and continue running
            #     self.message_update("Killing all excel", True)
            #     os.system("taskkill /f /t /im EXCEL.EXE")  # Kill all Excel files in background
            # if ret == 1:  # Back to previous step
            #     return

    def timestamp_name(self, input_, button):  # Define txt and xlsx report name by timestamp
        """ # Get timestamp """
        now = datetime.datetime.now()  # Get current day and time
        now = now.strftime("%Y%m%d_%H%M%S")  # 20250312_162924
        self.status_update('Step 1 of 7 : Clicked "Parsing FW report"')  # Update status
        self.message_update("Getting timestamp : " + now, True)
        file_name_excel = r"SWtestParsingReport\NoValidType_"  # Default file name if type_ is invalid

        if button == "FW":
            # Build Excel & txt name
            file_name_excel = r"SWtestParsingReport\\SWtestParsingReport_" + input_ + "_" + now + '.xlsx'
            self.file_name_txt = r"SWtestParsingReport\SWtestParsingReport_" + input_ + "_" + now + '.txt'
        elif button == "INI":
            # Build Excel & txt name
            file_name_excel = r"SWtestParsingReport" + "\\" + self.sub_items[0] + ".xlsx"
            self.file_name_txt = r"SWtestParsingReport" + "\\" + self.sub_items[0] + '.txt'
        elif button == "TOOL":
            file_name_excel = r"SWtestParsingReport\\SWtestParsingReport_TOOL_" + now + '.xlsx'
            self.file_name_txt = r"SWtestParsingReport\SWtestParsingReport_TOOL_" + now + '.txt'
        # File name ==> SWtestParsingReport_20250312_162924.xlsx(or txt)
        file_name_excel_abs = self.code_path + "\\" + file_name_excel  # Get complete path(absolute path)

        return file_name_excel, file_name_excel_abs

    def excel_build(self, file_name_excel):  # Create folder and build excel
        if self.excel_build_on:
            """ # Define/Generate folder to store xlsx report """
            __excel_handle = PythonExcel(self.template_path + "\\" + self.template)
            path = self.code_path + r"\SWtestParsingReport"
            if not os.path.isdir(path):  # If there is no directory "\SWtestParsingReport", create one
                os.mkdir(path)  # create "\SWtestParsingReport" directory
            # Copy template to "\SWtestParsingReport" directory

            if not os.path.isfile(r"SWtestParsingReport\\HIL.xlsx"):        # No HIL file = FW/INI case
                __excel_handle.copy_xlsx((self.template_path + "\\" + self.template), file_name_excel)
            else:                                                           # HIL case
                os.rename(r"SWtestParsingReport\\HIL.xlsx", file_name_excel)
                os.rename(r"SWtestParsingReport\\HIL.txt", self.file_name_txt)

            del __excel_handle
            time.sleep(self.delay_time)

    def path_searching(self, file_name, keyword):  # Transfer dataframe to final excel report
        """ # Examples """
        # Input : TQSTRYDTXR_2.9500 , 2.TestData\\
        # Output : [ D:\Coding\TVD\Data_path\S29_VCOL\2. Release\TQSTRYDTXR_2.9500\2.TestData\ ]

        # Input : TQSTRYDTXR_2.96 , 2.TestData\\
        # Output : [  D:\Coding\TVD\Data_path\S29_VCOL\2. Release\TQSTRYDTXR_2.9600\2.TestData\,
        #             D:\Coding\TVD\Data_path\S29_VCOL\2. Release\TQSTRYDTXR_2.9601\2.TestData\     ]

        return_list = []
        # self.message_update("Searching folders(sub folders) under " + file_name, True)
        # self.status_update("Searching folders(sub folders) under " + file_name)

        """ # Build path database """
        for i in range(len(self.path_list)):  # Run all items from list of data path
            if os.path.isdir(self.path_list[i]):  # if self.path_list[i] is a valid path, run ...
                list_folder = os.listdir(self.path_list[i])  # Get all items under data path
                for j in list_folder:
                    if file_name in str(j):  # If there is any folder that matches file_name
                        return_list.append(self.path_list[i] + "\\" + str(j) + "\\" + keyword)
        if len(return_list) == 0:  # If unable to return value in previous loop
            QMessageBox.warning(self, "Warning", "Fail : Cannot find " + str(file_name))
        return return_list

    def folder_parsing(self, button, folder_list, path_, type_, file_name_excel_abs, sub_path_list):  # Loop and scan
        bible_output = self.bible_output  # Define empty dataframe for bible list searching

        for h in range(len(path_)):     # Run all path
            # df_output = self.df_empty.copy(deep=True)
            # replace = True  # Decide df_output should be replaced or not. Once replaced, True => False
            self.message_update("Running folder : ", True)
            self.message_update(path_[h], False)

            if button == "FW" or button == "INI":       # FW, INI case
                """ # Prepare to build hex and ini row """
                self.message_update("Building hex and ini information", True)
                path_hex = path_[h].replace("2.TestData", "1.ImageFile")
                # ... 2. Release\TQSTRYDTXR_2.9600\2.TestData\ => ... 2. Release\TQSTRYDTXR_2.9600\1.ImageFile\
                path_ini = path_[h].replace("2.TestData", "3.Ini")
                # ... 2. Release\TQSTRYDTXR_2.9600\2.TestData\ => ... 2. Release\TQSTRYDTXR_2.9600\3.Ini\
                df_hex_ini = self.df_empty.copy(deep=True)

                # path_hex = self.path_searching(fw_ver_, "1.ImageFile\\")
                # path_ini = self.path_searching(fw_ver_, "3.Ini\\")
                # print("Path_hex : ", path_hex)
                # print("Path_ini : ", path_ini)

                """ # Get hex parameter """
                name_list_hex = [f for f in listdir(path_hex) if isfile(join(path_hex, f))]  # Get only file from folder
                hex_list = []  # hex_list = empty list for .hex file
                for i in range(len(name_list_hex)):
                    if ".hex" in name_list_hex[i]:  # If there is .hex in file name
                        hex_list.append(name_list_hex[i].split(".hex")[0])  # Remove ".hex" string and fill into hex_list
                hex_name = "NA"
                hex_name_path = "NA"
                if len(hex_list) != 0:  # If .hex file quantity is not 0
                    hex_name = hex_list[0]
                    hex_name_path = path_hex + hex_list[0] + ".hex"
                    for i in range(len(hex_list)):
                        if len(hex_name) > len(hex_list[i]) > 0:        # Get .hex file with the shortest name length
                            hex_name = hex_list[i]
                            hex_name_path = path_hex + hex_list[i] + ".hex"

                # print(hex_name, hex_name_path)
                # Fill into df_hex_ini dataframe
                df_hex_ini.iloc[0, 2] = hex_name
                self.Final_release_fw_ver = hex_name + ".hex"
                df_hex_ini.iloc[0, 21] = hex_name_path
                df_hex_ini.iloc[0, 10] = "1"
                df_hex_ini.iloc[0, 13] = "2400"
                df_hex_ini.iloc[0, 14] = "2400"
                df_hex_ini.iloc[0, 15] = "0"
                df_hex_ini.iloc[0, 16] = "Karl"

                """ # Get ini parameter """
                name_list_ini = os.listdir(path_ini)  # Get all file and folders from path_ini
                ini_list = []
                ini_list_path = []
                for g in range(len(name_list_ini)):
                    if ".ini" in name_list_ini[g]:
                        self.Final_release_ini_ver = name_list_ini[g]

                    if "_Rename_" in name_list_ini[g]:
                        self.RD_name = name_list_ini[g].split("_")[2]

                        # 1_Rename_rack_20240826_DTN => ...TQSTRYDTXR_2.9500\3.Ini\1_Rename_rack_20240826_DTN\
                        path_ini_full = path_ini + name_list_ini[g] + "\\"
                        # Get only file from folder
                        file_ini = [f for f in listdir(path_ini_full) if isfile(join(path_ini_full, f))]

                        for i in range(len(file_ini)):
                            if ".ini" in file_ini[i]:  # If there is .ini in file name
                                ini_list.append(
                                    file_ini[i].split(".ini")[0])  # Remove ".ini" string and fill into ini_list
                                ini_list_path.append(path_ini_full + file_ini[i])  # Fill full ini path into ini_list_path
                        # print(ini_list)
                        # print(ini_list_path)
                if len(ini_list) != 0:  # If .hex file quantity is not 0
                    for i in range(len(ini_list)):
                        # Add empty row to put parameter
                        df_hex_ini = pd.concat([df_hex_ini, self.df_empty], ignore_index=True)
                        df_hex_ini.iloc[i + 1, 3] = ini_list[i]
                        df_hex_ini.iloc[i + 1, 21] = ini_list_path[i]
                        df_hex_ini.iloc[i + 1, 10] = "1"
                        df_hex_ini.iloc[i + 1, 13] = "2400"
                        df_hex_ini.iloc[i + 1, 14] = "2400"
                        df_hex_ini.iloc[i + 1, 15] = "0"
                        df_hex_ini.iloc[i + 1, 16] = "Rack"
                df_output = df_hex_ini.copy(deep=True)  # Add df_hex_ini in front of df_output
            elif button == "TOOL":  # Tool case
                df_output = None        # Tool case has no ini/hex row

            for i in range(len(folder_list)):
                full_path = ""
                if button == "TOOL":        # Tool case
                    full_path = path_[h]
                elif button == "FW":            # FW case
                    full_path = path_[h] + folder_list[i] + "\\"
                elif button == "INI":           # INI case
                    full_path = path_[h] + "2.6.iniTestReport\\" + folder_list[i] + "\\"

                if os.path.isdir(full_path):  # If full_path is a directory
                    self.message_update("Running folder : " + full_path, True)
                    # Get file name in folder
                    name_list = []
                    walk_list = [f for f in os.walk(full_path) if (".xls" in str(f))]
                    for j in range(len(walk_list)):
                        for k in range(len(walk_list[j][-1])):
                            if str(walk_list[j][0]).endswith("\\"):
                                name_list.append(walk_list[j][0] + walk_list[j][-1][k])
                            else:
                                name_list.append(walk_list[j][0] + "\\" + walk_list[j][-1][k])

                    if len(name_list) != 0:
                        for j in range(len(name_list)):
                            # If target xlsx contains "-SWE" ".xls" and has no "~$"
                            if "-SWE" in name_list[j] and "~$" not in name_list[j] and ".xls" in name_list[j]:
                                # Ver folder (1/2) Sub folder (2/6) Excel (3/10)
                                self.status_update("Step 2 of 7 : Ver folder (" + str(h + 1) + "/" + str(len(path_)) +
                                                   ") Sub folder (" + str(i + 1) + "/" + str(len(folder_list)) +
                                                   ") Excel (" + str(j + 1) + "/" + str(len(name_list)) + ")")

                                # Output full path for xlsx
                                # print(name_list[j])   # ...2.TestData\2.1.Tessy\\S-SWE4-TESSY_TQS2QRDTXR2_9500.xlsm
                                self.message_update("Running Excel : " + name_list[j][1:-1], True)
                                # Fill info df_xlsx from xlsx path, keyword = Pass, fail, warning
                                df_xlsx = self.dataframe_fill(name_list[j], " PASS FAIL WARNING ", type_)

                                if len(df_xlsx) != 0:
                                    if df_output is None:       # Tool case has no ini/hex row
                                        df_output = df_xlsx
                                    else:
                                        df_output = pd.concat([df_output, df_xlsx], ignore_index=True)
                                        # print("Sheet : " + df_xlsx)
                                        # print("Excel : " + df_output)

                            """ # In INI case, after running all sub folder items, run sub path items """
                            if button == "INI" and j == len(name_list) - 1 \
                                    and i == len(folder_list) - 1 and h == len(path_) - 1:
                                for k in range(len(sub_path_list)):
                                    self.status_update("Step 3 of 7 : INI path : (" + str(k + 1) + "/" +
                                                       str(len(sub_path_list)) + ")")
                                    if os.path.isfile(sub_path_list[k]):  # If sub file exist
                                        self.message_update("Running Excel : " + sub_path_list[k][1:-1], True)
                                        df_xlsx = self.dataframe_fill(sub_path_list[k], " PASS FAIL WARNING ", type_)

                                        if len(df_xlsx) != 0:
                                            df_output = pd.concat([df_output, df_xlsx], ignore_index=True)
                                    else:
                                        self.message_update("Excel not existing : " + sub_path_list[k], True)

            """ # Count types from Customer, Redmine No """
            set_cus = set()  # Use "set" function to count different items without repeating them
            set_red = set()
            self.set_owner = set()
            self.set_reviewer = set()
            for i in range(len(df_output)):  # Collect types into sets
                set_cus.add(df_output.iloc[i, 0])
                set_red.add(df_output.iloc[i, 1])
                self.set_owner.add(df_output.iloc[i, 16])
                self.set_reviewer.add(df_output.iloc[i, 17])
            for i in ["", "NA", "nan", "None"]:  # Remove useless items
                set_cus.discard(i)
                set_red.discard(i)
                self.set_owner.discard(i)
                self.set_reviewer.discard(i)
            for i in ["Karl", "Rack"]:  # Remove RD from owner members
                self.set_owner.discard(i)

            # print(len(set_cus), "x set_cus : ", set_cus)
            # print(len(set_red), "x set_red : ", set_red)
            if len(set_cus) == 1:  # Fill customer into first row if customer is unique
                df_output.iloc[0, 0] = max(set_cus)
            # else:                           # If not unique, fill NA(And NA will be highlighted)
            #     df_output.iloc[0, 0] = "NA"
            if len(set_red) == 1:  # Fill redmine into first row if redmine is unique
                df_output.iloc[0, 1] = max(set_red)
            # else:                           # If not unique, fill NA(And NA will be highlighted)
            #     df_output.iloc[0, 1] = "NA"

            """ # Add two empty columns for parameters """
            df_output.insert(0, 'Ref_1', "", allow_duplicates=False)  # Add empty row for parameters
            df_output.insert(1, 'Ref_2', "", allow_duplicates=False)  # Add empty row for parameters

            """ # Transfer dataframe to Excel """
            if self.excel_build_on:
                self.message_update("Generating xlsx report, please wait for a while . . . . . .", True)
                self.status_update("Step 4 of 7 : Generating xlsx report . . .")
                # Transfer dataframe to final excel report
                # self.build_xlsx(df_output, file_name_excel, file_name_excel_abs)

                __excel_handle = PythonExcel(file_name_excel_abs)  # Refresh path and call win32com function again
                if len(bible_output) != 0:
                    __excel_handle.write_pandas(df_output, "Sheet1", 2 + len(bible_output), 1, True)
                    # self.build_xlsx(df_output, file_name_excel_abs, 2 + len(bible_output), 1)
                else:
                    __excel_handle.write_pandas(df_output, "Sheet1", 2, 1, True)
                    # self.build_xlsx(df_output, file_name_excel_abs, 2, 1)
                del __excel_handle
                time.sleep(self.delay_time)

            # Save df_output into bible_output
            for i in range(len(df_output)):
                # df_output.iloc[i, 4] : FW ver.
                # df_output.iloc[i, 5] : INI ver.
                # df_output.iloc[i, 9] : Test Case No
                # df_output.iloc[i, 16] : RWM
                array_ = [df_output.iloc[i, 4], df_output.iloc[i, 5], df_output.iloc[i, 9], df_output.iloc[i, 16]]
                bible_output.append(array_)

        return bible_output

    def html_parsing(self, path_, hil_customer, hil_redmine, hil_reviewer):
        # path_ = [r"\\10.3.0.101\Public\Software\3_FW_release_record\Hex\TQS\MB\1.QRUA\2.TQS3\1.TQS3\TQSDQRWCXE_5.2102"
        #          r"\2.TestData\2.6.iniTestReport\TQSDQRJMCCWA24X_R5_2102_Y_XXXXX_015\TestCaseRun_20241202_163652_946"
        #          r"\T-HILP-FMIXX-TQSDQRJMCCWA24X_R5_2102_Y_XXXXX_015.html",
        #          r"\\10.3.0.101\Public\Software\3_FW_release_record\Hex\TQS\MB\1.QRUA\2.TQS3\1.TQS3\TQSDQRWCXE_5.2102"
        #          r"\2.TestData\2.6.iniTestReport\TQSDQRJMCCWA24X_R5_2102_Y_XXXXX_015\TestCaseRun_20241129_162139_001"
        #          r"\T-HILP-FMIXX-TQSDQRJMCCWA24X_R5_2102_Y_XXXXX_015.html"]
        # hil_customer = "JMCCC"
        # hil_redmine = "#185477"
        # hil_reviewer = "Wuwuwuwuwu"
        summary_order_total = 0
        df_excel = self.df_empty.copy(deep=True)  # Empty dataframe for output(Final)
        df_sheet_copy = self.df_empty.copy(deep=True)  # Empty dataframe for sheets(Temp)

        """ # Loop all html path """
        for h in range(len(path_)):
            html_name = path_[h]
            self.message_update("Running HTML : ", True)
            self.message_update(html_name, False)
            self.status_update("Step 2 of 7 : HTML (" + str(h + 1) + "/" + str(len(path_)) + ")")

            """ # Read HTML info """
            env_list, sum_list = self.html_read(path_[h])

            """ # Fill in dataframe """
            summary_order_total = 0
            df_sheet = self.df_empty.copy(deep=True)  # Dataframe to record valid value
            summary_order = 0  # Calculate how many summary items in one sheet

            """ # Fill in Environment part """
            # Fixed items
            df_sheet.iloc[0, 10] = "1"      # Test Qty always be "1"
            df_sheet.iloc[0, 20] = "TVD1"   # Department always be "TVD1"
            df_sheet.iloc[0, 21] = html_name    # Fill in FilePath blank

            # Scanning items for the first summary row
            for i in range(0, 7):  # Loop first 7 test items from "self.test_items_keyword"
                for j in range(len(env_list)):  # Loop Environment part
                    # If Environment item is not empty
                    if env_list[j][1] != "None":
                        if "Case" in env_list[j][0]:  # If Environment title is "Test Case" or "Task Case"
                            df_sheet.iloc[0, 6] = env_list[j][1]
                        elif "工具軟體" in env_list[j][0]:  # If Environment title contains "工具軟體"
                            df_sheet.iloc[0, 19] = env_list[j][1]
                        elif "人員" in env_list[j][0]:  # If Environment title contains "人員"
                            # Input : Xiuwen(Suzy)    Output : Xiuwen
                            df_sheet.iloc[0, 16] = env_list[j][1].split("(")[0]
                            # Input : Xiuwen(Suzy)    Output : Suzy
                            df_sheet.iloc[0, 17] = env_list[j][1].split("(")[-1].split(")")[0]
                        elif "king" in env_list[j][0]:  # If Environment title contains "Actual_Working.Minute"
                            try:
                                rwm = float(env_list[j][1])  # Try to get minute
                                df_sheet.iloc[0, 14] = int(rwm)  # Minute to integer
                            except (TypeError, ValueError):  # If unable to get minute, give ""
                                df_sheet.iloc[0, 14] = ""
                        elif "硬體" in env_list[j][0]:  # If Environment title contains "硬體"
                            df_sheet.iloc[0, 4] = env_list[j][1]
                        # If Environment title contains test item key word
                        elif self.test_items_keyword[i] in env_list[j][0]:
                            df_sheet.iloc[0, i] = env_list[j][1]

            # Currently, df_sheet will get one line with repeated data

            """ # Fill in Summary part """
            if summary_order == 0:  # Initialize dataframe at first summary item
                df_sheet_copy = df_sheet.copy(deep=True)

            if len(sum_list) != 0:  # If able to detect "Results" from excel
                for i in range(len(sum_list)):  # Run loop from (Results +1) to (Results +100)
                    # If keyword PASS, FAIL, WARNING were found, and keyword is not " "
                    keywords = " PASS FAIL WARNING "
                    if sum_list[i][0] in keywords and len(str(sum_list[i][0])) > 3:
                        # If summary_order = 0, fill into dataframe
                        # If summary_order >=1, append dataframe before fill in data
                        if summary_order != 0:
                            df_sheet = pd.concat([df_sheet, df_sheet_copy], ignore_index=True)
                        # Count test result into "self.summary"
                        test_result = str(sum_list[i][0])
                        continue_ = True  # Message to decide we should record result or not
                        if "PASS" in test_result:
                            self.summary[0] = self.summary[0] + 1
                        elif "FAIL" in test_result:
                            self.summary[1] = self.summary[1] + 1
                        elif "WARNING" in test_result:
                            self.summary[2] = self.summary[2] + 1
                        else:
                            continue_ = False  # No valid result => Don't record result

                        if continue_:  # Valid result => Record result
                            for j in range(0, 23):  # Run first 20 test items (Environment items on excel)
                                df_sheet.iloc[summary_order, j] = df_sheet.iloc[0, j]
                            df_sheet.iloc[summary_order, 0] = hil_customer  # Record "customer"
                            df_sheet.iloc[summary_order, 1] = hil_redmine  # Record "redmine No"
                            df_sheet.iloc[summary_order, 7] = sum_list[i][1].split("_")[0]  # Record "Test Case No"
                            df_sheet.iloc[summary_order, 8] = sum_list[i][2]  # Record "Test Case"
                            df_sheet.iloc[summary_order, 9] = sum_list[i][3]  # Record "Test Item"
                            df_sheet.iloc[summary_order, 11] = sum_list[i][0]  # Record "Test Result"
                            df_sheet.iloc[summary_order, 17] = hil_reviewer      # Record "Reviewer"
                            df_sheet.iloc[summary_order, 18] = "HIL"  # Record "Project"

                            summary_order = summary_order + 1
                            summary_order_total = summary_order_total + 1
                # At the end of loop, df_sheet record to df_excel
                # RWM = Actual working minute(Environment) / summary qty
                for j in range(summary_order):  # Loop all summary items
                    rwm = df_sheet.iloc[j, 14]  # Get Actual working minute(Environment)
                    if rwm != "None" and rwm != "":  # If Actual working minute(Environment) exists
                        df_sheet.iloc[j, 14] = round(rwm / summary_order, 2)  # Get float till 2 digits

                if summary_order != 0:  # If valid summary item > 0, fill in df_excel
                    # print(df_excel.iloc[0, 0])
                    pd.set_option('display.max_columns', None)  # Switch to show all dataframe columns
                    if len(df_excel.iloc[0, 0]) == 0:  # If first time filling, df_excel = df_sheet
                        df_excel = df_sheet
                        # print(summary_order)
                    else:  # If not first time filling, combine them
                        df_excel = pd.concat([df_excel, df_sheet], ignore_index=True)
                        # print(summary_order)

        # print(df_excel.head(5))
        # pd.set_option('display.max_columns', None)  # Switch to show all dataframe columns
        # print(df_excel)
        time.sleep(self.delay_time)

        """ # Return result """
        if summary_order_total == 0 and len(df_excel) == 1:     # If Excel file is an empty report
            return ""
        else:
            return df_excel

    def para_col_build(self, bible_output, fw_ver_, type_, hexini_, file_name_excel_abs):
        """ # Count types from FW ver, ini ver """
        set_fw = set()  # Use "set" function to count different items without repeating them
        set_ini = set()
        for i in range(len(bible_output)):  # Collect types into sets
            set_fw.add(bible_output[i][0])  # FW ver.
            set_ini.add(bible_output[i][1])  # INI ver.
        for i in ["", "NA", "nan", "None"]:  # Remove useless items
            set_fw.discard(i)
            set_ini.discard(i)
        print("set_fw", set_fw)
        print("set_ini", set_ini)
        set_fw = set([i for i in set_fw if '.' not in str(i)])  # Only save items without "."
        set_ini = set([i for i in set_ini if '.' not in str(i)])  # Only save items without "."
        print("set_fw sort", set_fw)
        print("set_ini sort", set_ini)

        """ # Left side add two columns for parameters """
        # bible_output.insert(0, 'Ref_1', "", allow_duplicates=False)    # Add empty row for parameters
        # bible_output.insert(1, 'Ref_2', "", allow_duplicates=False)    # Add empty row for parameters
        # noinspection SpellCheckingInspection
        ref_1 = ["FW Ver:", "Type:", "hex_release_note.txt path:", "ini_release_note.txt path:", "Bible Ver:",
                 "Error message:", "", "", "當前parsing版本:", "當前parsing檔案:", "TestCaseCount:",
                 "PassCount:", "FailCount:", "hexConut:", "iniCount:", "", "INI:"]
        ref_2 = [fw_ver_, type_, hexini_, "", self.bible_list_full.split("\\")[-1], "", "", "", fw_ver_, "parsing list",
                 len(bible_output), self.summary[0] + self.summary[2], self.summary[1],
                 len(set_fw), len(set_ini), "", ""]
        # if len(bible_output) < len(ref_1):     # Fill empty rows into df_output to match parameter column length
        #     for i in range(len(ref_1) - len(bible_output)):
        #         bible_output = pd.concat([bible_output, self.df_empty], ignore_index=True)

        """ # Build data frame for extra parameter columns """
        parameter_col = ["ref_1", "ref_2"]

        # Fill in ini table items(User input
        for i in range(len(self.sub_list)):
            # ref_1.append(self.sub_list[i])
            ref_1.append("")
            ref_2.append("")

        # Define empty dataframe with headers
        df_parameter = pd.DataFrame(columns=parameter_col, index=list(range(0, len(ref_1))))
        for row in range(len(df_parameter)):  # Cleaning dataframe from "NAN" to ""
            df_parameter.iloc[row, 0] = ref_1[row]
            df_parameter.iloc[row, 1] = ref_2[row]

        __excel_handle = PythonExcel(file_name_excel_abs)
        __excel_handle.wrap_text(sheet=2, location="B4", wrap_on=True)
        __excel_handle.write_pandas(df_parameter, "Sheet1", 2, 1, False)
        del __excel_handle
        time.sleep(self.delay_time)

    def bible_list_build(self, df_output, type_, file_name_excel_abs):
        bible_path = self._bible
        bible_list = self.bible_list_get(bible_path)  # Get bible list

        """ # Save bible list as csv in local to release memory """
        __excel_handle = PythonExcel(bible_list)
        bible_list = __excel_handle.excel_to_csv(self.bible_list_csv, 1)  # BibleList.xlsx => BibleList.csv
        del __excel_handle
        # bible_list = r"C:\Users\jimmy_lee\Desktop\WinMergeCompare\SWTestList_V293.csv"  # Get bible list

        self.message_update("Scanning bible list, please wait for a while . . . . . .", True)
        self.status_update("Step 5 of 7 : Scanning bible list, please wait for a while . . . . . .")
        __excel_handle = PythonExcel(bible_list)
        df_bible = __excel_handle.sheet_to_df_by_no(bible_list, 1)  # Get first sheet from bible list
        # print("bible list df length", "data frame bible list")
        # print(len(df_bible))
        # print(df_bible)

        """ # Run all items from dataframe output """
        for i in range(0, len(df_output)):
            # df_output[i][0] : FW ver.
            # df_output[i][1] : INI ver.
            # df_output[i][2] : Test Case No
            # df_output[i][3] : RWM
            # df_output[i][4] : "EWM"  (Add from this stage)
            # df_output[i][5] : "Duration"  (Add from this stage)
            # df_output[i][6] : "Classification"  (Add from this stage)

            test_item_id = df_output[i][2]  # Get test item ID from each row in df_output

            # Tool case : Use special type to do searching
            if self.tool_type_ is not None:     # Tool case, self.tool_type_ != None, Other case, self.tool_type_ = None
                filter_project = ((df_bible[df_bible.columns[0]] == self.tool_type_) | (
                            df_bible[df_bible.columns[0]] == "HIL"))
            # If bible list "column A" = "Type" from UI
            # QRUA case : Allows "TQS", "TQS2", "TQS3", "TQSX", "TQSK", "TQSU", "MOH"
            elif type_ == "QRUA":
                or_list = []
                # noinspection SpellCheckingInspection
                qrua_list = ["TQS", "TQS2", "TQS3", "TQSX", "TQSK", "TQSU", "MOH", "HIL"]
                for j in range(len(qrua_list)):
                    or_list.append(df_bible[df_bible.columns[0]] == qrua_list[j])
                # Get all rows that matches any item from QRUA list
                filter_project = (or_list[0] | or_list[1] | or_list[2] | or_list[3] |
                                  or_list[4] | or_list[5] | or_list[6] | or_list[7])
            # Non QRUA case : Only allows "Type" from UI
            else:
                filter_project = ((df_bible[df_bible.columns[0]] == type_) | (df_bible[df_bible.columns[0]] == "HIL"))
            # Find bible list column 1 = test item ID
            filter_test_item = (df_bible[df_bible.columns[1]] == test_item_id)
            # Record bible list row that matches "project" and "test item ID" at the same time
            filter_result = df_bible[(filter_project & filter_test_item)]
            # print(filter_test_item)
            # print(filter_result)

            """ # Check how many row in bible list matches """
            # One row matching : fill in all items
            if len(filter_result) == 1:
                index_ = filter_result.index[0]  # Get index(Row) number from filter dataframe
                df_output[i].append(filter_result.loc[index_, 6])  # Record "EWM" as df_output[i][4]
                # Calculate "Duration"
                try:
                    # Record "Duration" as df_output[i][5] via EWM - RWM
                    df_output[i].append(math.ceil(float(df_output[i][4]) - float(df_output[i][3])))
                except (TypeError, ValueError):  # If unable to get minute, give ""
                    df_output[i].append("")

                df_output[i].append(filter_result.loc[index_, 8])  # Record "Classification" as df_output[i][6]
                if self.tool_type_ is not None:             # Tool case, project = "type"
                    df_output[i].append(type_)
                else:
                    df_output[i].append(filter_result.loc[index_, 0])  # Record "Project" as df_output[i][7]
                self.message_update("In bible list, successfully found " + type_ + " with " + test_item_id, False)
            # No row matching : Cannot find the item
            elif len(filter_result) == 0:
                self.message_update("In bible list, unable to find " + type_ + " with " + test_item_id, False)
            # Multiple rows matching : Bible list item repeating error
            else:
                self.message_update("In bible list, found more than one " + type_ + "with " + test_item_id, False)

        list_duration = df_output

        """ # Build data frame for extra parameter columns """
        self.message_update("Filling bible list items, please wait for a while . . . . . .", True)

        duration_col = ["ewm", "rwm", "duration"]
        classification_col = ["classification"]
        # Define empty dataframe with headers
        df_duration = pd.DataFrame(columns=duration_col, index=list(range(0, len(list_duration))))
        df_classification = pd.DataFrame(columns=classification_col, index=list(range(0, len(list_duration))))
        df_project = pd.DataFrame(columns=classification_col, index=list(range(0, len(list_duration))))

        for i in range(len(list_duration)):
            for row in range(len(df_duration)):  # Cleaning dataframe from "NAN" to ""
                if list_duration[row][3] == "2400":
                    df_duration.iloc[row, 0] = "2400"
                    df_duration.iloc[row, 1] = "2400"
                    df_duration.iloc[row, 2] = "0"
                    df_classification.iloc[row, 0] = ""
                    df_project.iloc[row, 0] = ""
                elif len(list_duration[row]) > 4:
                    df_duration.iloc[row, 0] = list_duration[row][4]
                    df_duration.iloc[row, 1] = list_duration[row][3]
                    df_duration.iloc[row, 2] = list_duration[row][5]
                    df_classification.iloc[row, 0] = list_duration[row][6]
                    df_project.iloc[row, 0] = list_duration[row][7]
                else:
                    df_duration.iloc[row, 0] = ""
                    df_duration.iloc[row, 1] = list_duration[row][3]
                    df_duration.iloc[row, 2] = ""
                    df_classification.iloc[row, 0] = ""
                    df_project.iloc[row, 0] = ""
        __excel_handle = PythonExcel(file_name_excel_abs)
        # __excel_handle.write_pandas(df_duration, "Sheet1", 2, 16, True)

        __excel_handle.write_pandas(df_classification, "Sheet1", 2, 8, True)
        __excel_handle.write_pandas(df_project, "Sheet1", 2, 21, True)
        __excel_handle.write_pandas(df_duration, "Sheet1", 2, 16, True)
        del __excel_handle
        time.sleep(self.delay_time)
        # self.build_xlsx(df_duration, file_name_excel_abs, 2, 16)

        # print(df_output)
        # return df_output

    def highlight(self, file_name_abs, row_end, col_end):
        # Highlight abnormal value into colors
        row_offset = 1
        col_offset = 1

        """ # Get data frame from xlsx report """
        __excel_handle = PythonExcel(file_name_abs)  # Refresh path and call win32com function again
        # Get dataframe
        df_output = __excel_handle.sheet_to_df_by_name('Sheet1', 1, 1, row_end, col_end)
        # df_output = __excel_handle.scan_sheet(file_name_abs)

        """ # Highlight abnormal value in sheet via dataframe """
        self.message_update("Highlighting abnormal values in TVD report...", True)

        __excel_handle.select_worksheet('Sheet1')
        # print(df_output)
        for i in range(1, len(df_output)):  # Loop row
            self.status_update("Step 6 of 7 : Highlighting abnormal values in TVD report (" + str(i + 1) + "/" +
                               str(len(df_output)) + ")")
            # print(df_d1.iloc[i])                  # Print full row
            for j in range(2, len(df_output.columns)):  # Loop column
                if df_output.iloc[i, j] == "":
                    # Paint into dark yellow
                    __excel_handle.fill_color(i + row_offset, j + col_offset, 150, 150, 50)
                else:
                    value_ = df_output.iloc[i, j]
                    try:
                        number_ = float(value_)  # Try to get number
                        if number_ < 0:  # Highlight negative numbers
                            # Paint into dark yellow
                            __excel_handle.fill_color(i + row_offset, j + col_offset, 150, 150, 50)
                        else:
                            __excel_handle.fill_color(i + row_offset, j + col_offset, 255, 255, 255)
                    except (TypeError, ValueError):  # If unable to get number
                        # Highlight NA and None items
                        if str(value_) == 'NA' or str(value_) == 'None':
                            # Paint into dark yellow

                            __excel_handle.fill_color(i + row_offset, j + col_offset, 150, 150, 50)
                        elif "." in str(value_):  # Highlight file extension(.hex or .ini)
                            if j == 4 or j == 5:  # Target column : "FW ver" and "ini ver"
                                # Paint into dark yellow
                                __excel_handle.fill_color(i + row_offset, j + col_offset, 150, 150, 50)
            if "FAIL" in str(df_output.iloc[i, 13]):  # Highlight fail result
                for j in range(2, len(df_output.columns)):  # Loop column
                    __excel_handle.fill_color(i + row_offset, j + col_offset, 255, 0, 0)  # Paint into blood-red
        del __excel_handle
        # time.sleep(self.delay_time)

    def pivot_analysis(self, file_name_excel_abs, last_row):
        """ # Build pivot table """
        self.message_update("Building pivot table at first page...", True)

        # If last_row = 100, get range from C1 to X100(Including title and blank)
        source_data = [2, "C1:X" + str(last_row + 2)]
        pivot_row = 76                  # Fixed value
        pivot_range = "B76:DE850"       # Fixed value

        __excel_handle = PythonExcel(file_name_excel_abs)  # Refresh path and call win32com function

        # FW test result - pass list
        __excel_handle.pivot_table(
            pivot_table_name="FW test result pass",
            data_from=source_data, pivot_to=[1, [2, pivot_row]],
            pivot_row=["FW ver.", "Task", "Test Case No"], pivot_column=["Test Result"],
            pivot_value=["FilePath"],
            filter_row=[["Task", ["ini", "FMI", "HIL"], False],
                        ["FW ver.", ["(blank)"], False],
                        ["Test Case No", ["(blank)"], False]],
            filter_column=["Test Result", ["PASS"]]
        )

        # FW test result - fail/warning list
        __excel_handle.pivot_table(
            pivot_table_name="FW test result fail",
            data_from=source_data, pivot_to=[1, [9, pivot_row]],
            pivot_row=["FW ver.", "Task", "Test Case No"], pivot_column=["Test Result"],
            pivot_value=["FilePath"],
            filter_row=[["Task", ["ini", "FMI", "HIL"], False],
                        ["FW ver.", ["(blank)"], False],
                        ["Test Case No", ["(blank)"], False]],
            filter_column=["Test Result", ["WARNING", "FAIL"]]
        )

        # hex fail-test item
        __excel_handle.pivot_table(
            pivot_table_name="HEX fail-test item",
            data_from=source_data, pivot_to=[1, [16, pivot_row]],
            pivot_row=["Task", "Test Case No", "FW ver."], pivot_column=["Test Result"],
            pivot_value=["FilePath"],
            filter_row=[["Task", ["ini", "FMI", "HIL"], False],
                        ["FW ver.", ["(blank)"], False],
                        ["Test Case No", ["(blank)"], False]],
            filter_column=["Test Result", ["FAIL"]]
        )

        # hex fail-test item(pass)
        fail_test_item_list = __excel_handle.pivot_read(
            pivot_sheet=1, pivot_table_name="HEX fail-test item", keyword="SWE"
        )
        __excel_handle.pivot_table(
            pivot_table_name="HEX fail-test item(pass)",
            data_from=source_data, pivot_to=[1, [23, pivot_row]],
            pivot_row=["Task", "Test Case No", "FW ver."], pivot_column=["Test Result"],
            pivot_value=["FilePath"],
            filter_row=[["Task", ["ini", "FMI", "HIL"], False],
                        ["Test Case No", fail_test_item_list, True],
                        ["FW ver.", ["(blank)"], False]],
            filter_column=["Test Result", ["PASS"]]
        )

        # hex warning-test item
        __excel_handle.pivot_table(
            pivot_table_name="HEX warning-test item",
            data_from=source_data, pivot_to=[1, [30, pivot_row]],
            pivot_row=["Task", "Test Case No", "FW ver."], pivot_column=["Test Result"],
            pivot_value=["FilePath"],
            filter_row=[["Task", ["ini", "FMI", "HIL"], False],
                        ["FW ver.", ["(blank)"], False],
                        ["Test Case No", ["(blank)"], False]],
            filter_column=["Test Result", ["WARNING"]]
        )

        # hex warning-test item(pass)
        fail_test_item_list = __excel_handle.pivot_read(
            pivot_sheet=1, pivot_table_name="hex warning-test item", keyword="SWE"
        )
        __excel_handle.pivot_table(
            pivot_table_name="hex warning-test item(pass)",
            data_from=source_data, pivot_to=[1, [37, pivot_row]],
            pivot_row=["FW ver.", "Task", "Test Case No"], pivot_column=["Test Result"],
            pivot_value=["FilePath"],
            filter_row=[["Task", ["ini", "FMI", "HIL"], False],
                        ["Test Case No", fail_test_item_list, True],
                        ["FW ver.", ["(blank)"], False]],
            filter_column=["Test Result", ["PASS"]]
        )

        # ---------------------------------------------
        # ini test result - pass list
        __excel_handle.pivot_table(
            pivot_table_name="INI test result pass",
            data_from=source_data, pivot_to=[1, [44, pivot_row]],
            pivot_row=["INI ver.", "Task", "Test Case No"], pivot_column=["Test Result"],
            pivot_value=["FilePath"],
            filter_row=[["Task", ["ini", "FMI", "HIL"], True],
                        ["INI ver.", ["(blank)"], False],
                        ["Test Case No", ["(blank)"], False]],
            filter_column=["Test Result", ["PASS"]]
        )

        # ini test result - fail/warning list
        __excel_handle.pivot_table(
            pivot_table_name="INI test result fail",
            data_from=source_data, pivot_to=[1, [51, pivot_row]],
            pivot_row=["INI ver.", "Task", "Test Case No"], pivot_column=["Test Result"],
            pivot_value=["FilePath"],
            filter_row=[["Task", ["ini", "FMI", "HIL"], True],
                        ["INI ver.", ["(blank)"], False],
                        ["Test Case No", ["(blank)"], False]],
            filter_column=["Test Result", ["WARNING", "FAIL"]]
        )

        # ini fail-test item
        __excel_handle.pivot_table(
            pivot_table_name="ini fail-test item",
            data_from=source_data, pivot_to=[1, [58, pivot_row]],
            pivot_row=["Task", "Test Case No", "INI ver."], pivot_column=["Test Result"],
            pivot_value=["FilePath"],
            filter_row=[["Task", ["ini", "FMI", "HIL"], True],
                        ["INI ver.", ["(blank)"], False],
                        ["Test Case No", ["(blank)"], False]],
            filter_column=["Test Result", ["FAIL"]]
        )

        # ini fail-test item(pass)
        fail_test_item_list = __excel_handle.pivot_read(
            pivot_sheet=1, pivot_table_name="ini fail-test item", keyword="SWE"
        )
        __excel_handle.pivot_table(
            pivot_table_name="ini fail-test item(pass)",
            data_from=source_data, pivot_to=[1, [65, pivot_row]],
            pivot_row=["INI ver.", "Task", "Test Case No"], pivot_column=["Test Result"],
            pivot_value=["FilePath"],
            filter_row=[["Task", ["ini", "FMI", "HIL"], True],
                        ["Test Case No", fail_test_item_list, True],
                        ["INI ver.", ["(blank)"], False]],
            filter_column=["Test Result", ["PASS"]]
        )

        # ini warning-test item
        __excel_handle.pivot_table(
            pivot_table_name="ini warning-test item",
            data_from=source_data, pivot_to=[1, [72, pivot_row]],
            pivot_row=["INI ver.", "Task", "Test Case No"],
            pivot_column=["Test Result"],
            pivot_value=["FilePath"],
            filter_row=[["Task", ["ini", "FMI", "HIL"], True],
                        ["INI ver.", ["(blank)"], False],
                        ["Test Case No", ["(blank)"], False]],
            filter_column=["Test Result", ["WARNING"]]
        )

        # ini warning-test item(pass)
        fail_test_item_list = __excel_handle.pivot_read(
            pivot_sheet=1, pivot_table_name="ini warning-test item", keyword="SWE"
        )
        __excel_handle.pivot_table(
            pivot_table_name="ini warning-test item(pass)",
            data_from=source_data, pivot_to=[1, [79, pivot_row]],
            pivot_row=["INI ver.", "Task", "Test Case No"], pivot_column=["Test Result"],
            pivot_value=["FilePath"],
            filter_row=[["Task", ["ini", "FMI", "HIL"], True],
                        ["Test Case No", fail_test_item_list, True],
                        ["INI ver.", ["(blank)"], False]],
            filter_column=["Test Result", ["PASS"]]
        )

        # Define pivot table color
        __excel_handle.fill_color_range(sheet=1, range=pivot_range, red=36, green=64, blue=98)
        # Define pivot table word color
        __excel_handle.fill_word_color_range(sheet=1, range=pivot_range, red=255, green=255, blue=255)
        # Define pivot table word type
        __excel_handle.fill_word_type_range(sheet=1, range=pivot_range, type="@")
        # Define pivot table font
        # noinspection SpellCheckingInspection
        __excel_handle.fill_font_range(sheet=1, range=pivot_range, font="Calibri")

    def hexini_change(self, file_name_excel_abs, release_note_link):
        """ # Add fw/ini release note """
        self.message_update("Adding fw/ini release note...", True)

        # release_note_link = "http://203.74.156.241:1314/redmine/projects/release-note/wiki/" \
        #                     "Request_2179_TQSDMRXXXE_5339;" \
        #                     "http://203.74.156.241:1314/redmine/projects/release-note/wiki/" \
        #                     "Request_2179_TQSDMRXXXE_539"
        # file_name_excel_abs = r"D:\Coding\TVD_tool\SWtestParsingReport\
        # SWtestParsingReport_TQSTRYDTXR_2.9500_20251023_161952 - 複製.xlsx"

        self.message_update("Generating hex/ini release note", True)
        os.system(self.tool_path + "\\Parsing_release_note.exe FW " + release_note_link)        # Run cmd, generate txt

        __excel_handle = PythonExcel(file_name_excel_abs)   # Refresh path and call win32com function
        df_change = self.df_empty_single.copy(deep=True)    # Initialize empty dataframe to edit Excel

        ini_row = 45        # ini starts from row 45
        hex_row = 41        # hex starts from row 41

        """ # Detect & fill in ini change part """
        ini_release_note = self.tool_path + "\\ini_release_note.txt"
        if os.path.isfile(ini_release_note):
            self.message_update("Ini release note : OK !", False)
            with open(ini_release_note, 'r', encoding="utf-8") as file:
                row = 0
                for line in file:
                    self.message_update(line.strip(), False)
                    df_change.iloc[0, 0] = line.strip()                     # Feed ini into dataframe
                    __excel_handle.write_pandas(df_change, 1, ini_row + row, 2, False)  # Feed dataframe into Excel
                    df_change.iloc[0, 0] = "Test ID:"
                    __excel_handle.write_pandas(df_change, 1, ini_row + row, 11, False)  # Feed "Test ID:" into Excel
                    __excel_handle.write_pandas(df_change, 1, ini_row + row, 15, False)  # Feed "Test ID:" into Excel
                    row = row + 1
                    __excel_handle.insert_row(1, ini_row + row)                     # Insert row
                file.close()
        else:
            self.message_update("Unable to get ini release note", False)

        """ # Detect & fill in hex change part """
        hex_release_note = self.tool_path + "\\hex_release_note.txt"
        if os.path.isfile(hex_release_note):
            self.message_update("Hex release note : OK !", False)
            with open(hex_release_note, 'r', encoding="utf-8") as file:      # Open txt
                row = 0
                for line in file:
                    self.message_update(line.strip(), False)
                    df_change.iloc[0, 0] = line.strip()                     # Feed ini into dataframe
                    __excel_handle.write_pandas(df_change, 1, hex_row + row, 2, False)  # Feed dataframe into Excel
                    df_change.iloc[0, 0] = "Test ID:"
                    __excel_handle.write_pandas(df_change, 1, hex_row + row, 11, False)  # Feed "Test ID:" into Excel
                    __excel_handle.write_pandas(df_change, 1, hex_row + row, 15, False)  # Feed "Test ID:" into Excel
                    row = row + 1
                    __excel_handle.insert_row(1, hex_row + row)                     # Insert row
                file.close()
        else:
            self.message_update("Unable to get hex release note", False)

    def sorting_page_arrange(self, file_name_excel_abs):
        """ # Filling sorting page """
        self.message_update("Filling sorting page...", True)

        __excel_handle = PythonExcel(file_name_excel_abs)   # Refresh path and call win32com function
        df_change = self.df_empty_single.copy(deep=True)    # Initialize empty dataframe to edit Excel

        """ Fill in top left part in the first sheet """
        # Get timestamp
        now = datetime.datetime.now()  # Get current day and time
        df_change.iloc[0, 0] = now.strftime("%Y/%m/%d")    # Get current day and time (2025/03/12)
        __excel_handle.write_pandas(df_change, 1, 5, 3, False)

        # FW version
        df_change.iloc[0, 0] = self.Final_release_fw_ver
        __excel_handle.write_pandas(df_change, 1, 6, 3, False)

        # INI version
        df_change.iloc[0, 0] = self.Final_release_ini_ver
        __excel_handle.write_pandas(df_change, 1, 7, 3, False)

        df_change.iloc[0, 0] = "-- spec ver --"
        __excel_handle.write_pandas(df_change, 1, 8, 3, False)

        # RD name
        df_change.iloc[0, 0] = self.RD_name
        __excel_handle.write_pandas(df_change, 1, 9, 3, False)

        df_change.iloc[0, 0] = ""                                        # Tasker
        __excel_handle.write_pandas(df_change, 1, 10, 3, False)

        # Tester and reviewer
        tester = ""
        reviewer = ""
        for i in self.set_owner:
            tester = tester + ", " + i
        tester = tester[2:]             # Remove ", " in the beginning
        for i in self.set_reviewer:
            reviewer = reviewer + ", " + i
        reviewer = reviewer[2:]         # Remove ", " in the beginning

        df_change.iloc[0, 0] = tester
        __excel_handle.write_pandas(df_change, 1, 11, 3, False)

        df_change.iloc[0, 0] = reviewer
        __excel_handle.write_pandas(df_change, 1, 12, 3, False)

        df_change.iloc[0, 0] = self.TVD_manager
        __excel_handle.write_pandas(df_change, 1, 13, 3, False)
        df_change.iloc[0, 0] = "-- Bug ID --"
        __excel_handle.write_pandas(df_change, 1, 14, 3, False)

        del __excel_handle

    def screenshot_run(self, file_name_excel_abs):
        """ # Getting screenshot """
        self.message_update("Getting screenshot...", True)

        # git_tag = "HCS1CABOXX_1.1213"
        git_tag = self.git_tag

        __excel_handle = PythonExcel(file_name_excel_abs)  # Refresh path and call win32com function
        png = os.popen(self.tool_path + "\\get_GitTag_img.exe " + git_tag).read().strip()
        png_path = self.tool_path + "\\" + png
        target = "AC5:AR18"
        __excel_handle.select_worksheet("Sorting")
        # __excel_handle.insert_image_to_range(png_path, target, lock_ratio=True, img_width=0, img_height=0)
        __excel_handle.insert_image_to_range(png_path, target, lock_ratio=True)

    # Parsing ini related functions
    def signal_receiver(self, bool_, list_, list2_):
        self.sub_signal = bool_
        self.sub_list = list_
        self.sub_items = list2_

    # self.folder_parsing functions
    def dataframe_fill(self, excel_full_path, keywords, type_):  # Transfer Excel path to dataframe
        # excel_full_path = r"D:\Coding\TVD_tool\T-SWEX-VCXXX_VCOL-XXX-A-XX-RE-XXXX-01.02.xlsm"
        # keywords = " PASS FAIL WARNING "

        summary_order_total = 0         #

        """ # Generate empty dataframe """
        df_excel = self.df_empty.copy(deep=True)        # Dataframe for output(Final)
        df_sheet_copy = self.df_empty.copy(deep=True)   # Dataframe for sheets(Temp)

        """ # Get xlsx sheets """
        __excel_handle = PythonExcel(excel_full_path)               # Call win32com function
        self.message_update("Excel sheets :", False)
        sheet_list = __excel_handle.scan_sheet(excel_full_path)     # Read all sheet name
        self.message_update(str(sheet_list), False)                 # Show all sheet name

        """ # Run sheets with keyword "Environment (Precondition)" and "工具軟體版本" """
        for h in sheet_list:                                    # Run all sheets
            sheet_name = h
            sheet_key = __excel_handle.read_value(h, ["B4"])    # Read "B4" box
            sheet_key_2 = __excel_handle.read_value(h, ["B5"])  # Read "B5" box
            pd.set_option('display.max_columns', None)          # Switch to show all dataframe columns

            # If sheet has key word : Environment (Precondition)
            if str(sheet_key) in "Environment (Precondition)" and str(sheet_key_2) == "工具軟體版本":
                self.message_update("Running target sheet : " + sheet_name, False)
                df = __excel_handle.sheet_to_df_by_name(sheet_name, 1, 1, 150, 7)  # Raw dataframe, from A1 to G100
                df_sheet = self.df_empty.copy(deep=True)        # Dataframe to record valid value
                summary_order = 0  # Calculate how many summary items in one sheet

                """ # Fill in Environment part """
                # Fixed items
                df_sheet.iloc[0, 10] = "1"           # Test Qty always be "1"
                df_sheet.iloc[0, 20] = "TVD1"        # Department always be "TVD1"
                df_sheet.iloc[0, 21] = str(excel_full_path)  # Fill in FilePath blank

                # Scanning items for the first summary row
                for i in range(0, 7):           # Loop first 7 test items from "self.test_items_keyword"
                    for j in range(0, 20):          # Loop Excel Environment part : row 0 to 19
                        # If Environment item is not empty (df[3] means excel column "D")
                        if df[3][j] != "None":
                            if "Case" in str(df[1][j]):  # If Environment title is "Test Case" or "Task Case"
                                df_sheet.iloc[0, 6] = str(df[3][j])
                            elif "測報" in str(df[1][j]):  # If Environment title contains "測報"
                                df_sheet.iloc[0, 19] = str(df[3][j])
                            elif "人員" in str(df[1][j]):  # If Environment title contains "人員"
                                # Input : Xiuwen(Suzy)    Output : Xiuwen
                                df_sheet.iloc[0, 16] = str(df[3][j]).split("(")[0]
                                # Input : Xiuwen(Suzy)    Output : Suzy
                                if "(" in str(df[3][j]):
                                    df_sheet.iloc[0, 17] = str(df[3][j]).split("(")[-1].split(")")[0]
                            elif "king" in str(df[1][j]):  # If Environment title contains "Actual_Working.Minute"
                                try:
                                    rwm = float(df[3][j])  # Try to get minute
                                    df_sheet.iloc[0, 14] = int(rwm)  # Minute to integer
                                except (TypeError, ValueError):  # If unable to get minute, give ""
                                    df_sheet.iloc[0, 14] = ""
                            # If Environment title contains test item key word
                            elif self.test_items_keyword[i] in str(df[1][j]):
                                df_sheet.iloc[0, i] = str(df[3][j])

                # Currently, df_sheet will get one line with repeated data

                """ # Fill in Summary part """
                if summary_order == 0:      # Initialize dataframe at first summary item
                    df_sheet_copy = df_sheet.copy(deep=True)

                summary_start = 0           # This is where to start scanning summary items
                for i in range(20, 60):     # Find keyword "Results" to start counting summary
                    if str(df[1][i]) in "Results" and len(str(df[1][i])) > 3:  # Scanning at B column(df[1])
                        summary_start = i + 1
                        break

                if summary_start != 0:      # If able to detect "Results" from excel
                    for i in range(summary_start, summary_start + 150):  # Run loop from (Results +1) to (Results +100)
                        # If keyword PASS, FAIL, WARNING were found, and keyword is not " "
                        if str(df[1][i]) in keywords and len(str(df[1][i])) > 3:
                            # If summary_order = 0, fill into dataframe
                            # If summary_order >=1, append dataframe before fill in data
                            if summary_order != 0:
                                df_sheet = pd.concat([df_sheet, df_sheet_copy], ignore_index=True)
                            # Count test result into "self.summary"
                            test_result = str(df[1][i])
                            continue_ = True  # Message to decide we should record result or not
                            if "PASS" in test_result:
                                self.summary[0] = self.summary[0] + 1
                            elif "FAIL" in test_result:
                                self.summary[1] = self.summary[1] + 1
                            elif "WARNING" in test_result:
                                self.summary[2] = self.summary[2] + 1
                            else:
                                continue_ = False  # No valid result => Don't record result

                            if continue_:  # Valid result => Record result
                                for j in range(0, 23):  # Run first 20 test items (Environment items on excel)
                                    df_sheet.iloc[summary_order, j] = df_sheet.iloc[0, j]
                                    # Copy environment item as previous value
                                df_sheet.iloc[summary_order, 7] = str(df[2][i])  # Record "Test Case No"
                                df_sheet.iloc[summary_order, 8] = str(df[3][i])  # Record "Test Case"
                                df_sheet.iloc[summary_order, 9] = str(df[5][i])  # Record "Test Item"
                                df_sheet.iloc[summary_order, 11] = str(df[1][i])  # Record "Test Result"

                                df_sheet.iloc[summary_order, 18] = type_  # Record "Project"
                                # print(i, str(df[1][i]), str(df[2][i]), str(df[3][i]), str(df[5][i]))
                                summary_order = summary_order + 1
                                summary_order_total = summary_order_total + 1

                        # End at line 'Test process record', df_sheet record to df_excel
                        elif 'Test process record' in str(df[1][i]) or 'Test process record' in str(df[2][i]) \
                                or 'Test process record' in str(df[3][i]):
                            # RWM = Actual working minute(Environment) / summary qty
                            for j in range(summary_order):  # Loop all summary items
                                rwm = df_sheet.iloc[j, 14]   # Get Actual working minute(Environment)
                                if rwm != "None" and rwm != "":  # If Actual working minute(Environment) exists
                                    df_sheet.iloc[j, 14] = round(rwm / summary_order, 2)  # Get float till 2 digits

                            if summary_order != 0:  # If valid summary item > 0, fill in df_excel
                                # print(df_excel.iloc[0, 0])
                                pd.set_option('display.max_columns', None)  # Switch to show all dataframe columns
                                if len(df_excel.iloc[0, 0]) == 0:  # If first time filling, df_excel = df_sheet
                                    df_excel = df_sheet
                                    # print(summary_order)
                                else:  # If not first time filling, combine them
                                    df_excel = pd.concat([df_excel, df_sheet], ignore_index=True)
                                    # print(summary_order)

                            break
                # print(df_excel.head(5))
                # pd.set_option('display.max_columns', None)  # Switch to show all dataframe columns
                # print(df_excel)
        del __excel_handle  # Close Excel
        time.sleep(self.delay_time)

        # If Excel file is an empty report
        if summary_order_total == 0 and len(df_excel) == 1:
            return ""
        else:
            return df_excel

    # Old functions
    def build_xlsx(self, df_output, file_name_abs, row_offset, col_offset):
        # Transfer dataframe to final excel report
        # Highlight abnormal value into colors

        """ # Generate xlsx report """
        __excel_handle = PythonExcel(file_name_abs)  # Refresh path and call win32com function again
        # Write dataframe value into template
        __excel_handle.write_pandas(df_output, "Sheet1", row_offset, col_offset, True)
        # __excel_handle.write_pandas(df_output, file_name_abs, "Sheet1", 2, 1, True)

        """ # Highlight abnormal value in sheet via output dataframe """
        self.message_update("Highlighting abnormal values in TVD report...", True)
        self.status_update("Highlighting abnormal values in TVD report...")
        __excel_handle.select_worksheet('Sheet1')
        for i in range(0, len(df_output)):  # Loop row
            # print(df_d1.iloc[i])                  # Show full row
            for j in range(len(df_output.columns)):  # Loop column
                if df_output.iloc[i, j] == "" and j > 1:
                    # Paint into dark yellow
                    __excel_handle.fill_color(i + row_offset, j + col_offset, 150, 150, 50)
                elif j > 1:
                    value_ = df_output.iloc[i, j]
                    try:
                        number_ = float(value_)  # Try to get number
                        if number_ < 0:  # Highlight negative numbers
                            # Paint into dark yellow
                            __excel_handle.fill_color(i + row_offset, j + col_offset, 150, 150, 50)
                        else:
                            __excel_handle.fill_color(i + row_offset, j + col_offset, 255, 255, 255)
                    except (TypeError, ValueError):  # If unable to get number
                        # Highlight NA and None items
                        if str(value_) == 'NA' or str(value_) == 'None':
                            # Paint into dark yellow
                            __excel_handle.fill_color(i + row_offset, j + col_offset, 150, 150, 50)
                        elif "." in str(value_):  # Highlight file extension(.hex or .ini)
                            if j == 4 or j == 5:  # Target column : "FW ver" and "ini ver"
                                # Paint into dark yellow
                                __excel_handle.fill_color(i + row_offset, j + col_offset, 150, 150, 50)
                        elif "FAIL" in str(value_):  # Highlight fail result
                            for k in range(3, len(df_output.columns) - 3):  # Loop column
                                __excel_handle.fill_color(i + row_offset, k, 255, 0, 0)  # Paint into blood-red
                        else:
                            __excel_handle.fill_color(i + row_offset, j + col_offset, 255, 255, 255)
        del __excel_handle
        time.sleep(self.delay_time)

    @staticmethod
    def folder_scan(root, target):
        print("root : ", root, "target : ", target, "\n")  # Show parameters
        count = 0  # Count searching times
        start_time = time.time()  # Set timer start
        for res in os.walk(root):  # os.walk function
            count = count + 1  # Add count
            # print(count, res)
            if target in str(res):  # If keyword was found
                print("Found folder : " + res[0] + "\\" + target)
                count = 0  # Initialize count if key word was found
                break
        if count > 0:  # If count was not initialized
            print("\n Cannot find " + target + "from " + root)
        end_time = time.time()  # Set timer end
        print("Time taken : " + str(start_time - end_time) + " seconds")  # Show searching time

    @staticmethod
    def html_read(html):
        # noinspection SpellCheckingInspection
        # html = r"\\10.3.0.101\Public\Software\3_FW_release_record\Hex\TQS\MB\1.QRUA\2.TQS3\1.TQS3" \
        #        r"\TQSDQRWCXE_5.2102\2.TestData\2.6.iniTestReport\TQSDQRJMCCWA24X_R5_2102_Y_XXXXX_015" \
        #        r"\TestCaseRun_20241202_163652_946\T-HILP-FMIXX-TQSDQRJMCCWA24X_R5_2102_Y_XXXXX_015.html"

        output_environment = []
        output_summary = []

        """ # Get HTML content until 'Test process record' """
        html_content = []
        with open(html, 'r', encoding="utf-8") as file:     # Read .html content
            for line in file:
                html_content.append(line)
                # print(line)
                if "Test process record" in line:       # Read until "Test process record" appears
                    break
            file.close()

        """ # Html content => line """
        html_by_line = []
        list_ = []
        for i in range(len(html_content)):
            list_.append(html_content[i])
            if "</tr>" in html_content[i]:          # Record list and clear it if html go to another line
                html_by_line.append(list_)
                list_ = []

        """ # Get environment box items """
        list_ = []
        for i in range(len(html_by_line)):                  # Run by html line
            # print("<<<<<<<<<<<<<<   >>>>>>>>>>>>>")
            for j in range(len(html_by_line[i])):           # Run by html box
                head = 'class="Text'
                if head in html_by_line[i][j]:              # Get summary item
                    list_.append(html_by_line[i][j].split('">')[-1].split("</td>")[0])
            if len(list_) > 0:
                output_environment.append(list_)
            list_ = []

        """ # Get summary box items """
        list_ = []
        for i in range(len(html_by_line)):                  # Run by html line
            # print("<<<<<<<<<<<<<<   >>>>>>>>>>>>>")
            for j in range(len(html_by_line[i])):           # Run by html box
                head = '<td  class="NumberCell'
                if head in html_by_line[i][j]:              # Get summary item
                    # print(line[j].split('">')[-1].split("</td>")[0])
                    list_.append(html_by_line[i][j].split('">')[-1].split("</td>")[0])
            if len(list_) > 0:
                output_summary.append(list_)
            list_ = []

        # print("output environment : ")
        # for i in range(len(output_environment)):
        #     print(output_environment[i])
        #
        # print("output summary : ")
        # for i in range(len(output_summary)):
        #     print(output_summary[i])

        return output_environment, output_summary


class SubWindowINI(QtWidgets.QDialog, Ui_Dialog_INI):
    # Define sub window signal type
    dialog_signal = QtCore.pyqtSignal(bool, list, list)

    def __init__(self):
        super(SubWindowINI, self).__init__()
        self.setupUi(self)
        self.table_path.setHorizontalScrollBarPolicy(QtCore.Qt.ScrollBarAlwaysOn)
        self.table_path.setVerticalScrollBarPolicy(QtCore.Qt.ScrollBarAlwaysOn)

        """ # Define default file name by timestamp """
        now = datetime.datetime.now()  # Get current day and time
        now = now.strftime("%Y%m%d_%H%M%S")  # 20250312_162924
        file_name_excel = r"SWtestParsingReport_" + now  # Build Excel name
        self.my_line_edit.setText(file_name_excel)

    @pyqtSlot()
    def on_button_continue_clicked(self):
        path_list = []  # Define empty list for output
        table_path_row = self.table_path.rowCount()  # Get row quantity from table

        """ # Get all readings from table """
        for i in range(table_path_row):
            if self.table_path.item(i, 0) is not None:  # If value from table is valid
                self.table_path.item(i, 0).setBackground(QtGui.QColor(255, 255, 255))  # Initialize row color
                if len(self.table_path.item(i, 0).text()) > 0:  # If value is not empty
                    path_list.append(self.table_path.item(i, 0).text())  # Record value into output list
                else:  # Record empty if value is empty
                    path_list.append("")

        """ # Checking invalid path from table """
        path_not_exist = False  # path all exist : Go to next step. Not all : Highlight and user enter again.
        path_list_exist = []  # Record existing path

        print("INI parsing items : ")
        for i in range(len(path_list)):  # Check all input from sub window
            print(path_list[i])
            if len(path_list[i].split("\\")) == 1:  # If item does not contain "\", record to folder list
                if len(path_list[i]) > 0:  # If value is not empty
                    path_list_exist.append(path_list[i])
            else:  # If contains "\" and existing, record to path list
                if not os.path.isfile(path_list[i]):
                    path_not_exist = True
                    # Show warning if path not existing
                    self.table_path.item(i, 0).setBackground(QtGui.QColor(255, 0, 0))  # Change row color
                else:
                    path_list_exist.append(path_list[i])

        if path_not_exist:
            QMessageBox.warning(self, "Warning !", 'Path not existing !')

        else:
            item_list = [self.my_line_edit.text()]
            if len(path_list) > 0:  # If path_list has something, deliver True + path_list
                self.dialog_signal.emit(True, path_list_exist, item_list)
            else:  # If path_list has nothing, deliver False + empty list
                self.dialog_signal.emit(False, [], item_list)

            self.close()  # Close sub window

    @pyqtSlot()
    def on_button_quit_clicked(self):
        # """ # Function 1 : Initialize table and close sub window """
        # self.table_path.clearContents()
        # self.table_path.update()
        # self.close()

        """ # Function 2 : Fill in default items """
        self.table_path.clearContents()
        # noinspection SpellCheckingInspection
        self.table_path.setItem(0, 0, QtWidgets.QTableWidgetItem(str("TQS2QRYUCCWA24X_R2_7306_Y_012D0_B01")))
        # noinspection SpellCheckingInspection
        self.table_path.setItem(0, 1, QtWidgets.QTableWidgetItem(str("TQS2QRYUCCWA24X_R2_7306_Y_018D0_B01")))
        # noinspection SpellCheckingInspection
        path = r"\\10.3.0.101\Public\Software\3_FW_release_record\Hex\TQS\MB\1.QRUA\TQS2QRYUXR_2.7305\2.TestData\2.1" \
               r".Tessy\S-SWE4-TESSY_TQS2QRYUXR2_7305.xlsm"
        # noinspection SpellCheckingInspection
        path1 = r"\\10.3.0.101\Public\Software\3_FW_release_record\Hex\TQS\MB\1.QRUA\TQS2QRYUXR_2.7305\2.TestData\2.1" \
                r".Tessy\8787.xlsm"
        # noinspection SpellCheckingInspection
        path2 = r"\\10.3.0.101\Public\Software\3_FW_release_record\Hex\TQS\MB\1.QRUA\TQS2QRYUXR_2.7305\2.TestData\2.1" \
                r".Tessy\7878.xlsm"
        self.table_path.setItem(0, 2, QtWidgets.QTableWidgetItem(str(path)))
        # self.table_path.setItem(0, 3, QtWidgets.QTableWidgetItem(str(path1)))
        # self.table_path.setItem(0, 4, QtWidgets.QTableWidgetItem(str(path2)))
        self.table_path.update()


class SubWindowHIL(QtWidgets.QDialog, Ui_Dialog_HIL):
    # Define sub window signal type
    dialog_signal = QtCore.pyqtSignal(bool, list, list)

    def __init__(self):
        super(SubWindowHIL, self).__init__()
        self.setupUi(self)
        # self.table_path.setHorizontalScrollBarPolicy(QtCore.Qt.ScrollBarAlwaysOn)
        # self.table_path.setVerticalScrollBarPolicy(QtCore.Qt.ScrollBarAlwaysOn)

        """ # Define default file name by timestamp """
        now = datetime.datetime.now()  # Get current day and time
        now = now.strftime("%Y%m%d_%H%M%S")  # 20250312_162924
        file_name_excel = r"SWtestParsingReport_HIL_" + now  # Build Excel name
        self.my_line_edit_4.setText(file_name_excel)

    @pyqtSlot()
    def on_button_fw_clicked(self):
        """ # Warning if any box is empty """
        if len(str(self.my_line_edit.text())) < 1:
            QMessageBox.warning(self, "Warning", '"Customer" can not be empty')
            return
        if len(str(self.my_line_edit_2.text())) < 1:
            QMessageBox.warning(self, "Warning", '"Redmine No" can not be empty')
            return
        if len(str(self.my_line_edit_3.text())) < 1:
            QMessageBox.warning(self, "Warning", '"Reviewer" can not be empty')
            return
        if len(str(self.my_line_edit_4.text())) < 1:
            QMessageBox.warning(self, "Warning", '"Report name" can not be empty')
            return

        """ # Check table path """
        path_list = []  # Define empty list for output
        table_path_row = self.table_path.rowCount()  # Get row quantity from table

        """ # Get all readings from table """
        for i in range(table_path_row):
            if self.table_path.item(i, 0) is not None:  # If value from table is valid
                self.table_path.item(i, 0).setBackground(QtGui.QColor(255, 255, 255))  # Initialize row color
                if len(self.table_path.item(i, 0).text()) > 0:  # If value is not empty
                    path_list.append(self.table_path.item(i, 0).text())  # Record value into output list
                else:  # Record empty if value is empty
                    path_list.append("")

        """ # Checking invalid path from table """
        path_not_exist = False  # path all exist : Go to next step. Not all : Highlight and user enter again.
        path_list_exist = []  # Record existing path

        print("HIL parsing items : ")
        for i in range(len(path_list)):  # Check all input from sub window
            print(path_list[i])
            if not os.path.isfile(path_list[i]) and path_list[i] != "":
                path_not_exist = True
                # Show warning if path not existing
                self.table_path.item(i, 0).setBackground(QtGui.QColor(255, 0, 0))  # Change row color
            else:
                path_list_exist.append(path_list[i])

        if path_not_exist:
            QMessageBox.warning(self, "Warning !", 'Path not existing !')

        else:
            item_list = [self.my_line_edit.text(), self.my_line_edit_2.text(),
                         self.my_line_edit_3.text(), self.my_line_edit_4.text(), "FW"]
            if len(path_list) > 0:  # If path_list has something, deliver True + path_list
                self.dialog_signal.emit(True, path_list_exist, item_list)
            else:  # If path_list has nothing, deliver False + empty list
                self.dialog_signal.emit(False, [], item_list)

            self.close()  # Close sub window

    @pyqtSlot()
    def on_button_ini_clicked(self):
        """ # Warning if any box is empty """
        if len(str(self.my_line_edit.text())) < 1:
            QMessageBox.warning(self, "Warning", '"Customer" can not be empty')
            return
        if len(str(self.my_line_edit_2.text())) < 1:
            QMessageBox.warning(self, "Warning", '"Redmine No" can not be empty')
            return
        if len(str(self.my_line_edit_3.text())) < 1:
            QMessageBox.warning(self, "Warning", '"Reviewer" can not be empty')
            return
        if len(str(self.my_line_edit_4.text())) < 1:
            QMessageBox.warning(self, "Warning", '"Report name" can not be empty')
            return

        """ # Check table path """
        path_list = []  # Define empty list for output
        table_path_row = self.table_path.rowCount()  # Get row quantity from table

        """ # Get all readings from table """
        for i in range(table_path_row):
            if self.table_path.item(i, 0) is not None:  # If value from table is valid
                self.table_path.item(i, 0).setBackground(QtGui.QColor(255, 255, 255))  # Initialize row color
                if len(self.table_path.item(i, 0).text()) > 0:  # If value is not empty
                    path_list.append(self.table_path.item(i, 0).text())  # Record value into output list
                else:  # Record empty if value is empty
                    path_list.append("")

        """ # Checking invalid path from table """
        path_not_exist = False  # path all exist : Go to next step. Not all : Highlight and user enter again.
        path_list_exist = []  # Record existing path

        print("HIL parsing items : ")
        for i in range(len(path_list)):  # Check all input from sub window
            print(path_list[i])
            if not os.path.isfile(path_list[i]) and path_list[i] != "":
                path_not_exist = True
                # Show warning if path not existing
                self.table_path.item(i, 0).setBackground(QtGui.QColor(255, 0, 0))  # Change row color
            else:
                path_list_exist.append(path_list[i])

        if path_not_exist:
            QMessageBox.warning(self, "Warning !", 'Path not existing !')

        else:
            item_list = [self.my_line_edit.text(), self.my_line_edit_2.text(),
                         self.my_line_edit_3.text(), self.my_line_edit_4.text(), "INI"]
            if len(path_list) > 0:  # If path_list has something, deliver True + path_list
                self.dialog_signal.emit(True, path_list_exist, item_list)
            else:  # If path_list has nothing, deliver False + empty list
                self.dialog_signal.emit(False, [], item_list)

            self.close()  # Close sub window

    @pyqtSlot()
    def on_button_quit_clicked(self):
        """ # Function 1 : Initialize table and close sub window """
        # self.table_path.clearContents()
        # self.table_path.update()
        # self.close()

        """ # Function 2 : Fill in default items """
        self.table_path.clearContents()
        # noinspection SpellCheckingInspection
        path = r"\\10.3.0.101\Public\Software\3_FW_release_record\Hex\TQS\MB\1.QRUA\2.TQS3\1.TQS3\TQSDQRWCXE_5.2102" \
               r"\2.TestData\2.6.iniTestReport\TQSDQRJMCCWA24X_R5_2102_Y_XXXXX_015\TestCaseRun_20241202_163652_946" \
               r"\T-HILP-FMIXX-TQSDQRJMCCWA24X_R5_2102_Y_XXXXX_015.html"
        # noinspection SpellCheckingInspection
        path1 = r"\\10.3.0.101\Public\Software\3_FW_release_record\Hex\TQS\MB\1.QRUA\2.TQS3\1.TQS3\TQSDQRWCXE_5.2102" \
                r"\2.TestData\2.6.iniTestReport\TQSDQRJMCCWA24X_R5_2102_Y_XXXXX_015\TestCaseRun_20241129_162139_001" \
                r"\T-HILP-FMIXX-TQSDQRJMCCWA24X_R5_2102_Y_XXXXX_015.html"

        self.table_path.setItem(0, 0, QtWidgets.QTableWidgetItem(str(path)))
        self.table_path.setItem(0, 1, QtWidgets.QTableWidgetItem(str(path1)))
        # noinspection SpellCheckingInspection
        # self.table_path.setItem(0, 2, QtWidgets.QTableWidgetItem(str("TQS2QRYUCCWA24X_R2_7306_Y_012D0_B01")))
        self.table_path.setItem(0, 3, QtWidgets.QTableWidgetItem(str("")))
        self.table_path.update()

        self.my_line_edit.setText("JMC")
        self.my_line_edit_2.setText("#1854")
        # noinspection SpellCheckingInspection
        self.my_line_edit_3.setText("Wuwu")


# =================[Main]====================
if __name__ == '__main__':
    app = QApplication(sys.argv)
    main_window = MainWindow()  # Run MainWindow initial
    if main_window.version_check():
        main_window.show()
        sys.exit(app.exec_())

    # main_window.folder_scan(r"D:\Coding\TVD", "TQS2QRDTXR_2.9500")
    # main_window.folder_scan(r"\\10.3.0.101\Public\Software\14_AutoTest\2_Tool\18_TVD_parsing_tool", "2.TestData")
    # key = "TQS2QRDTXR_2.9500"
    # main_window.folder_scan(r"\\10.3.0.101\Public\Software\3_FW_release_record\Hex", "TQS2QRDTXR_2.9500")
    # main_window.html_read()



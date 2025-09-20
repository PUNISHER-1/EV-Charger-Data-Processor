import pandas as pd 
import json
import sys
from PyQt5.QtWidgets import QFileDialog
import os
from process_data import DataThread
from PyQt5 import QtCore, QtGui, QtWidgets

class Ui_MainWindow(object):
    def setupUi(self, MainWindow):
        MainWindow.setObjectName("MainWindow")
        MainWindow.resize(895, 800)
        self.centralwidget = QtWidgets.QWidget(MainWindow)
        self.centralwidget.setObjectName("centralwidget")

        self.datafieldbox = QtWidgets.QScrollArea(self.centralwidget)
        self.datafieldbox.setGeometry(QtCore.QRect(10, 320, 831, 431))
        self.datafieldbox.setWidgetResizable(True)
        self.datafieldbox.setObjectName("datafieldbox")
        self.scrollAreaWidgetContents = QtWidgets.QWidget()
        self.scrollAreaWidgetContents.setGeometry(QtCore.QRect(0, 0, 829, 429))
        self.scrollAreaWidgetContents.setObjectName("scrollAreaWidgetContents")
        self.verticalLayout = QtWidgets.QVBoxLayout(self.scrollAreaWidgetContents)
        self.verticalLayout.setObjectName("verticalLayout")
        self.labelExample = QtWidgets.QLabel(self.scrollAreaWidgetContents)
        
        self.labelExample.setText("")
        self.labelExample.setObjectName("labelExample")
        self.verticalLayout.addWidget(self.labelExample)
        self.datafieldbox.setWidget(self.scrollAreaWidgetContents)
        self.calendarWidget = QtWidgets.QCalendarWidget(self.centralwidget)
        self.calendarWidget.setGeometry(QtCore.QRect(510, 60, 328, 197))
        self.calendarWidget.setObjectName("calendarWidget")
        self.verticalLayoutWidget = QtWidgets.QWidget(self.centralwidget)
        self.verticalLayoutWidget.setGeometry(QtCore.QRect(10, 60, 491, 101))
        self.verticalLayoutWidget.setObjectName("verticalLayoutWidget")
        self.verticalLayout1 = QtWidgets.QVBoxLayout(self.verticalLayoutWidget)
        self.verticalLayout1.setContentsMargins(0, 0, 0, 0)
        self.verticalLayout1.setObjectName("verticalLayout1")
        self.horizontalLayout = QtWidgets.QHBoxLayout()
        self.horizontalLayout.setObjectName("horizontalLayout")
        self.browsebox = QtWidgets.QLineEdit(self.verticalLayoutWidget)
        self.browsebox.setObjectName("browsebox")
        self.horizontalLayout.addWidget(self.browsebox)
        self.browse_btn = QtWidgets.QPushButton(self.verticalLayoutWidget)
        self.browse_btn.setObjectName("browse_btn")
        self.horizontalLayout.addWidget(self.browse_btn)
        self.verticalLayout1.addLayout(self.horizontalLayout)
        self.dropdownbrowse = QtWidgets.QComboBox(self.verticalLayoutWidget)
        self.dropdownbrowse.setObjectName("dropdownbrowse")
        self.verticalLayout1.addWidget(self.dropdownbrowse)
        self.path_lable = QtWidgets.QLabel(self.centralwidget)
        self.path_lable.setGeometry(QtCore.QRect(10, 30, 111, 16))
        self.path_lable.setObjectName("path_lable")
        self.date_lable = QtWidgets.QLabel(self.centralwidget)
        self.date_lable.setGeometry(QtCore.QRect(510, 30, 101, 20))
        self.date_lable.setObjectName("date_lable")
        self.data_lable = QtWidgets.QLabel(self.centralwidget)
        self.data_lable.setGeometry(QtCore.QRect(10, 290, 71, 16))
        self.data_lable.setObjectName("data_lable")
        self.generate_btn = QtWidgets.QPushButton(self.centralwidget)
        self.generate_btn.setGeometry(QtCore.QRect(10, 210, 91, 31))
        self.generate_btn.setObjectName("generate_btn")
        MainWindow.setCentralWidget(self.centralwidget)
        self.menubar = QtWidgets.QMenuBar(MainWindow)
        self.menubar.setGeometry(QtCore.QRect(0, 0, 895, 23))
        self.menubar.setObjectName("menubar")
        MainWindow.setMenuBar(self.menubar)
        self.statusbar = QtWidgets.QStatusBar(MainWindow)
        self.statusbar.setObjectName("statusbar")
        MainWindow.setStatusBar(self.statusbar)

        self.retranslateUi(MainWindow)
        QtCore.QMetaObject.connectSlotsByName(MainWindow)

        self.load_paths()
        self.dropdownbrowse.currentIndexChanged.connect(self.update_browsebox)

    def retranslateUi(self, MainWindow):
        _translate = QtCore.QCoreApplication.translate
        MainWindow.setWindowTitle(_translate("MainWindow", "MainWindow"))
        self.browse_btn.setText(_translate("MainWindow", "BROWSE"))
        self.browse_btn.clicked.connect(self.browse_folder)

        self.path_lable.setText(_translate("MainWindow", "PATH SELECTION"))
        self.date_lable.setText(_translate("MainWindow", "CHOOSE DATE"))
        self.data_lable.setText(_translate("MainWindow", "DATA FIELD"))
        self.generate_btn.setText(_translate("MainWindow", "GENERATE"))
        self.generate_btn.clicked.connect(self.generate)

    def browse_folder(self):
        # Open dialog box 
        folder_path,_ = QFileDialog.getOpenFileName(None, "Select File")
        
        if folder_path:
            # Set the selected folder path in the QLineEdit
            self.browsebox.setText(folder_path)
            
            if folder_path not in [self.dropdownbrowse.itemText(i) for i in range(self.dropdownbrowse.count())]:
                self.dropdownbrowse.addItem(folder_path)

            self.save_paths()

    def update_browsebox(self):
        selected_path = self.dropdownbrowse.currentText()
        self.browsebox.setText(selected_path) 
        
    def load_paths(self):
        if os.path.exists("paths.txt"):
            with open("paths.txt", "r") as file:
                for line in file:
                    path = line.strip()
                    if path:  
                        self.dropdownbrowse.addItem(path)

        if self.dropdownbrowse.count() > 0:
            self.dropdownbrowse.setCurrentIndex(0)
            self.update_browsebox()
            
    def save_paths(self):
        with open("paths.txt", "w") as file:
            for i in range(self.dropdownbrowse.count()):
                file.write(self.dropdownbrowse.itemText(i) + "\n")

    def generate(self):
        selected_date = self.calendarWidget.selectedDate()
        date_str = selected_date.toString("dd-MM-yyyy")

        folder_path = self.browsebox.text()

        extracted_data = extract_data_by_date_and_serial(excel_file=folder_path, selected_date=date_str)

        if extracted_data:
            self.data_worker = DataThread(extracted_data)
            self.data_worker.progress.connect(self.data_display)
            self.data_worker.finished.connect(self.cleanup_thread)
            self.data_worker.execute()
            # process_extracted_data(extracted_data, self)
            # print(json.dumps(extracted_data, indent=4, default=str))

    def cleanup_thread(self):
        self.data_worker.quit()
        self.data_worker.wait()  # Wait for the thread to finish
        self.data_worker = None  # set to None for garbage collection

    def data_display(self, message):
        new_label = QtWidgets.QLabel(message)
        self.verticalLayout.addWidget(new_label)

def extract_data_by_date_and_serial(excel_file, selected_date):
    df = pd.read_excel(excel_file)

    # Convert "Tested Date" column to datetime 
    df['Tested Date'] = pd.to_datetime(df['Tested Date'], format='%d-%m-%Y', errors='coerce')

    # Convert selected_date to a datetime object
    selected_date = pd.to_datetime(selected_date, format='%d-%m-%Y', errors='coerce')

    # Filter rows
    filtered_df = df[df['Tested Date'] == selected_date]

    # If no data
    if filtered_df.empty:
        print(f"No data found for the date: {selected_date.date()}")
        return None

    # unique models for the selected date
    unique_models = filtered_df['MODEL'].unique()

    models_list = []

    # Loop through each unique model
    for model in unique_models:
        model_df = filtered_df[filtered_df['MODEL'] == model]  # Filter rows for this specific model

        serials_dict = {}

        # Loop through each serial number in the current model
        for _, row in model_df.iterrows():
            serial_number = row['Charger Serial No']
            
            serials_dict[serial_number] = {
                "MODEL": row['MODEL'],
                "Charger Serial No": row['Charger Serial No'],
                "Gun A DCEM": row['Gun A DCEM'],
                "Gun B DCEM": row['Gun B DCEM'],
                "Gun C EM": row['Gun C EM'],
                "Upper sw ver.": row['Upper sw ver.'],
                "Pilot Cont sw": row['Pilot Cont sw'],
                "Tested Date": row['Tested Date'].strftime('%d-%m-%Y'),
                "Tested By": row['Tested BY'] 
            }
        
        models_list.append({
            "Model": model,
            "Serial Numbers": serials_dict
        })

    return models_list 


if __name__ == "__main__":
    import sys
    app = QtWidgets.QApplication(sys.argv)
    MainWindow = QtWidgets.QMainWindow()
    ui = Ui_MainWindow()
    ui.setupUi(MainWindow)
    MainWindow.show()
    sys.exit(app.exec_())

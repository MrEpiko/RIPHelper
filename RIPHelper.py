from genericpath import isfile
import sys
from PyQt6.QtWidgets import QApplication, QWidget, QDialog, QLabel, QDialogButtonBox, QVBoxLayout, QFileDialog, QComboBox, QMessageBox, QFrame
from PyQt6.QtGui import QIcon
from PyQt6 import uic
from PySide6.QtCore import QFileSystemWatcher
import json
from pathlib import Path
import subprocess
from datetime import datetime
import psutil
import winshell
import os
import win32com
import win32com.client
from win32event import CreateMutex
from win32api import GetLastError
from winerror import ERROR_ALREADY_EXISTS

class RemoveFolderDialog(QDialog):

    def __init__(self, folders_list):

        super().__init__()
        self.setWindowTitle("Remove folder")
        self.setWindowIcon(QIcon(f"./assets/icon.png"))     

        QBtn = QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel

        self.buttonBox = QDialogButtonBox(QBtn)
        self.buttonBox.accepted.connect(self.accept)
        self.buttonBox.rejected.connect(self.reject)

        self.layout = QVBoxLayout()
        message = QLabel("Which folder would you like to remove?")
        self.layout.addWidget(message)

        self.comboBox = QComboBox()
        for folder in folders_list:
            self.comboBox.addItem(folder)

        self.layout.addWidget(self.comboBox)
        self.layout.addWidget(self.buttonBox)
        self.setLayout(self.layout)


class MyApp(QWidget):

    def __init__(self):

        super().__init__()
        
        self.ui = uic.loadUi(f'./ui/gui.ui', self)             
        self.setFixedSize(self.size())
        
        self.setWindowIcon(QIcon(f"./assets/icon.png"))

        # Setting configurations
        with open(f"./json/config.json", "r+") as f:
            self.config = json.load(f) 
            f.close()
        
        self.folders_list = self.config["configurations"]["folders_list"]
        
        is_configured = self.config["is_configured"]
        if not is_configured:

            self.path_field.setText("Not configured...")
            self.path_field.setText("Not configured...")

            self.folders_box.setText("Not configured...")
            self.backups_box.setText("Not configured...")
            self.logs_box.setText("No logs to show.")

        else:

            frequency_value = self.config["configurations"]["frequency"]
            self.scans_frequency_box.setValue(frequency_value)

            execution_path = self.config["configurations"]["execution_program_path"]
            self.path_field.setText(execution_path)

            output = ""
            backup_output = ""
            for folder in self.folders_list:
                output += f"► {folder} \n"
                backup = self.folders_list[folder]["backup"]
                backup_output += f"► {backup} \n"
            
            self.folders_box.setText(output)
            self.backups_box.setText(backup_output)
        
        # Syncing scan status
        self.is_started = self.config["is_started"]
        if self.is_started:

            self.start_stop_button.setText("Stop")
            self.disabled_view.raise_()
            self.disabled_label.raise_()
            self.clear_scanned_files_button.lower()
            self.clear_logs_button.lower()

        # Setting watchers
        self.watcher = QFileSystemWatcher()
        self.watcher.addPath(f"./logs/logs.txt")
        self.watcher.fileChanged.connect(self.logs_update)

        self.files_watcher = QFileSystemWatcher()
        self.files_watcher.addPath(f"./json/scanned_files.json")
        self.files_watcher.fileChanged.connect(self.files_update)
        
        # Settings logs and scanned files
        with open(f"./logs/logs.txt", "r") as f:
            self.logs_box.setText(f.read())
            f.close()
        
        with open(f"./json/scanned_files.json", "r") as f:
            files_json = json.load(f)
        scanned_files = files_json["files"]
        
        output = ""
        for file in scanned_files:

            output += f"► {file} \n"

            old = scanned_files[file]["old_path"]
            old = old.replace("\\", "/")
            output += f"Old path: {old} \n"

            new = scanned_files[file]["new_path"]
            output += f"New path: {new} \n"
        
        self.scanned_files_box.setText(output)

        # Setting up events
        self.ui.start_stop_button.clicked.connect(self.start_stop_button_click)
        self.ui.browse_button.clicked.connect(self.browse_button_dialog)
        self.ui.add_folder_button.clicked.connect(self.add_folder_button_dialog)
        self.ui.remove_folder_button.clicked.connect(self.remove_folder_button_dialog)
        self.ui.save_button.clicked.connect(self.save_button_click)
        self.ui.clear_button.clicked.connect(self.clear_button_click)
        self.ui.clear_scanned_files_button.clicked.connect(self.clear_scanned_files_button_click)
        self.ui.clear_logs_button.clicked.connect(self.clear_logs_button_click)

    def logs_update(self):
        
        
        f = self.watcher.files()[0]
        with open(f, "r") as f:
            self.logs_box.setText(f.read())
            f.close()
        
        self.logs_box.verticalScrollBar().setValue(self.logs_box.verticalScrollBar().maximum())
    
    def files_update(self):

        f = self.files_watcher.files()[0]
        with open(f, "r") as f:
            files_json = json.load(f)
        
        scanned_files = files_json["files"]
        
        output = ""
        for file in scanned_files:

            output += f"► {file} \n"

            old = scanned_files[file]["old_path"]
            old = old.replace("\\", "/")
            output += f"Old path: {old} \n"

            new = scanned_files[file]["new_path"]
            output += f"New path: {new} \n"
        
        self.scanned_files_box.setText(output)

    
    def start_stop_button_click(self):

        button = self.start_stop_button
        current_text = button.text()

        is_configured = self.config["is_configured"]
        if not is_configured:

            message_box = QMessageBox.critical(
                self,
                "Error occured",
                "Program hasn't been configured yet, configure it first and try again!",
                buttons = QMessageBox.StandardButton.Ok, 

            )
            
            return
        
        date = datetime.now().strftime("[%H:%M:%S]")
            
        if current_text == "Start":

            is_running = False
            
            for proc in psutil.process_iter():
                
                try:
                    if "RIPHelperScan".lower() in proc.name().lower():
                        is_running = True
                except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
                    pass
            
            if is_running:

                QMessageBox.critical(
                    self,
                    "RIPHelper",
                    "You already have a scan process running or it's being killed at the moment. Try again in a few seconds.",
                    buttons = QMessageBox.StandardButton.Ok, 
                )
                return

            button.setText("Stop")
            self.config["is_started"] = True
            
            with open(f"./json/config.json", "w") as fw:
                json.dump(self.config, fw, indent=4)
                fw.close()  
            
            subprocess.Popen(f"./RIPHelperScan.exe", creationflags=subprocess.DETACHED_PROCESS | subprocess.CREATE_NEW_PROCESS_GROUP)
            self.disabled_view.raise_()
            self.disabled_label.raise_()
            self.clear_scanned_files_button.lower()
            self.clear_logs_button.lower()

            with open(f"./logs/logs.txt", "a") as f:
                f.write(f"{date} Scan started. \n")
                f.close()

            QMessageBox.information(
                self,
                "RIPHelper",
                "Scan has been successfully started!",
                buttons = QMessageBox.StandardButton.Ok, 
            )

        else:

            button.setText("Start")
            self.config["is_started"] = False

            with open(f"./json/config.json", "w") as fw:
                json.dump(self.config, fw, indent=4)
                fw.close()

            self.disabled_view.lower()
            self.disabled_label.lower()
            
            self.clear_scanned_files_button.raise_()
            self.clear_logs_button.raise_()

            with open(f"./logs/logs.txt", "a") as f:
                f.write(f"{date} Scan stop requested. \n")
                f.close()

            QMessageBox.information(
                self,
                "RIPHelper",
                "Scan has been successfully stopped!",
                buttons = QMessageBox.StandardButton.Ok, 
            )
    
    def browse_button_dialog(self):

        home_dir = str(Path.home())
        fname = QFileDialog.getOpenFileName(self, 'Select file', home_dir, filter = "Executable files (*.exe)")

        if fname[0]:

            self.path_field.setText(fname[0])
    
    def add_folder_button_dialog(self):

        home_dir = str(Path.home())
        fname = QFileDialog.getExistingDirectory(self, 'Select folder', home_dir)

        if fname:

            folder = fname
            if folder in self.folders_list:
                return
            
            fname_2 = QFileDialog.getExistingDirectory(self, 'Select backup folder', home_dir)

            if fname_2:

                self.folders_list[folder] = {"backup": fname_2}
                
                output = ""
                backup_output = ""
                for folder in self.folders_list:
                    output += f"► {folder} \n"
                    backup = self.folders_list[folder]["backup"]
                    backup_output += f"► {backup} \n"
                
                self.folders_box.setText(output)
                self.backups_box.setText(backup_output)
    
    def remove_folder_button_dialog(self):

        if len(self.folders_list) == 0:
            return

        dlg = RemoveFolderDialog(self.folders_list)
        button = dlg.exec()
        
        if button == 1:
            
            choice = dlg.comboBox.currentText()
            del self.folders_list[choice]
            
            output = ""
            backup_output = ""
            for folder in self.folders_list:
                output += f"► {folder} \n"
                backup = self.folders_list[folder]["backup"]
                backup_output += f"► {backup} \n"
            
            if output == "":
                output = "Not configured..."
                backup_output = "Not configured..."
                
            self.folders_box.setText(output)
            self.backups_box.setText(backup_output)
    
    def clear_button_click(self):

        message_box = QMessageBox.question(
            self,
            "Restart configurations",
            "Are you sure that you want to entirely restart all program's configurations? Doing so will also stop the current scan process from running.",
            buttons = QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No, 
        )

        if message_box == QMessageBox.StandardButton.No:
            return
        
        self.start_stop_button.setText("Start")

        self.scans_frequency_box.setValue(5)
        self.path_field.setText("Not configured...")
        self.folders_box.setText("Not configured...")
        self.backups_box.setText("Not configured...")

        self.folders_list = {}
        with open(f"./json/config.json", "r+") as f:
            self.config = json.load(f) 
            f.close()
        
        data = {

            "is_configured": False,
            "is_started": False,
            "configurations": {
                "frequency": 0,
                "execution_program_path": "",
                "folders_list": {}
            }

        }
        
        with open(f"./json/config.json", "w") as fw:
            json.dump(data, fw, indent=4)
            fw.close()
        
        with open(f"./json/config.json", "r+") as f:
            self.config = json.load(f) 
            f.close()
        
        QMessageBox.information(
            self,
            "Action successful",
            "Successfully restarted configurations!",
            buttons = QMessageBox.StandardButton.Ok, 
        )

    
    def save_button_click(self):
        
        number_value = self.scans_frequency_box.value()
        if (number_value < self.scans_frequency_box.minimum()) or (number_value > self.scans_frequency_box.maximum()):

            QMessageBox.critical(
                self,
                "Error occured",
                "The frequency number you have set is invalid, change it and try again!",
                buttons = QMessageBox.StandardButton.Ok, 

            )
            return
        
        execution_program_value = self.path_field.text()
        if execution_program_value == "Not configured...":
            QMessageBox.critical(
                self,
                "Error occured",
                "You did not set up the execution program path!",
                buttons = QMessageBox.StandardButton.Ok, 

            )
            return
        
        if "exiftool" not in execution_program_value.lower():
            QMessageBox.critical(
                self,
                "Error occured",
                "Program path that you have provided doesn't seem to be a path to Exiftool.",
                buttons = QMessageBox.StandardButton.Ok, 

            )
            return
        
        if len(self.folders_list) == 0:
            QMessageBox.critical(
                self,
                "Error occured",
                "You did not configure folders that should be scanned!",
                buttons = QMessageBox.StandardButton.Ok, 

            )
            return
        
        message_box = QMessageBox.question(
            self,
            "Save configurations",
            "Are you sure that you want to save new set of configurations and entirely override the old one?",
            buttons = QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No, 
        )

        if message_box == QMessageBox.StandardButton.No:
            return
        
        data = {

            "is_configured": True,
            "is_started": False,
            "configurations": {
                "frequency": self.scans_frequency_box.value(),
                "execution_program_path": self.path_field.text(),
                "folders_list": self.folders_list
            }

        }

        with open(f"./json/config.json", "w") as fw:
            json.dump(data, fw, indent=4)
            fw.close()
        
        with open(f"./json/config.json", "r+") as f:
            self.config = json.load(f) 
            f.close()

        QMessageBox.information(
            self,
            "Action successful",
            "Successfully saved configurations!",
            buttons = QMessageBox.StandardButton.Ok, 
        )
    
    def clear_logs_button_click(self):

        message_box = QMessageBox.question(
            self,
            "Clear logs",
            "Are you sure that you want to entirely clear the logs file? This action cannot be undone.",
            buttons = QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No, 
        )

        if message_box == QMessageBox.StandardButton.No:
            return
        
        with open(f"./logs/logs.txt", "w+") as f:
            f.truncate(0)
            f.close()
        
        message_box = QMessageBox.information(
            self,
            "Action successful",
            "Logs file has been successfully cleared.",
            buttons = QMessageBox.StandardButton.Ok, 
        )
    
    def clear_scanned_files_button_click(self):
        
        message_box = QMessageBox.question(
            self,
            "Clear scanned files",
            "Are you sure that you want to entirely clear scanned files? Doing so will make the scan go through the same files again and, therefore, it may cause the issues. This action cannot be undone.",
            buttons = QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No, 
        )

        if message_box == QMessageBox.StandardButton.No:
            return
        
        data = {

            "files": {}

        }
        
        with open(f"./json/scanned_files.json", "w+") as fw:
            json.dump(data, fw, indent=4)
        
        message_box = QMessageBox.information(
            self,
            "Action successful",
            "Scanned files have been successfully cleared.",
            buttons = QMessageBox.StandardButton.Ok, 
        )
        

if __name__ == "__main__":

    handle = CreateMutex(None, 1, 'A unique mutex name')

    if GetLastError(  ) == ERROR_ALREADY_EXISTS:
        
        pass

    else:

        app = QApplication(sys.argv)
        window = MyApp()
        window.show()

        window.logs_box.verticalScrollBar().setValue(window.logs_box.verticalScrollBar().maximum())

        startup = winshell.startup()

        path = os.path.join(startup, 'RIPHelperScan.lnk')
        shell = win32com.client.Dispatch("WScript.Shell")
        shortcut = shell.CreateShortCut(path)

        path = os.path.dirname(os.path.realpath(__file__))
        path = path.replace("\\", "/")

        shortcut.Targetpath = f"{path}/RIPHelperScan.exe" 
        shortcut.IconLocation = f"{path}/assets/favicon.ico"

        shortcut.save()    

        sys.exit(app.exec())
    

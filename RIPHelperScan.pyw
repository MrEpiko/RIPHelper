import json
import time
import pathlib
import plyer
from datetime import datetime
import shutil
import subprocess

class Scanner():

    def __init__(self):
        
        super().__init__()

        with open(f"./json/config.json", "r+") as f:
            self.config = json.load(f) 
            f.close()

        is_configured = self.config["is_configured"]
        if not is_configured:
            return
        
        self.frequency = self.config["configurations"]["frequency"]
        self.execution_program_path = self.config["configurations"]["execution_program_path"]
        self.folders_list = self.config["configurations"]["folders_list"]

        self.is_started = self.config["is_started"]

        with open(f"./json/scanned_files.json") as f:
            self.scanned_files_json = json.load(f)
            f.close()
        
        self.checked_files = self.scanned_files_json["files"]
    
    def add_to_startup(self):
        
        pass    
    
    def scan(self):

        while True:

            with open(f"./json/config.json", "r") as f:
                self.config = json.load(f) 
                f.close()
            self.is_started = self.config["is_started"]

            if not self.is_started:
                
                output = ""
                date = datetime.now().strftime("[%H:%M:%S]")
                output += f"{date} Scan stopped. \n"
                with open(f"./logs/logs.txt", "a") as f:
                    f.write(output)
                    f.close()
                break
    
            output = ""
            date = datetime.now().strftime("[%H:%M:%S]")
            
            output += f"{date} Scan started, folders: {self.folders_list} \n"
            
            traces = 0
            
            for folder in self.folders_list:

                backup_folder = self.folders_list[folder]["backup"]
                output += f"{date} Scanning {folder}... \n"

                p = pathlib.Path(folder)
                for f in p.rglob("*.tif"):
                    
                    stringed = str(f.name)
                    stringed_location = str(f)

                    if stringed in self.checked_files:
                        continue
                    
                    traces += 1
                    self.checked_files[stringed] = {"old_path": str(f), "new_path": f"{backup_folder}/{f.name}"}    
                    
                    try:
                        subprocess.call(f'{self.execution_program_path} -make=Creo "{stringed_location}" -overwrite_original', shell = False, creationflags = subprocess.CREATE_NO_WINDOW)
                    except:
                        output += f"{date} [{f.name}] There was an issue trying to run the external program. Please check your configuration and try again."
                        continue

                    time.sleep(1)

                    try:
                        shutil.move(f, backup_folder)
                    except shutil.Error:
                        pass

                    output += f"{date} Found .tif file: {f.name}. File has been executed and moved to a backup folder. \n"

            if traces > 0:

                output += f"{date} Scan finished, {traces} traces have been found! \n"
                
                self.scanned_files_json["files"] = self.checked_files
                
                with open(f"./json/scanned_files.json", "w") as fw:
                    json.dump(self.scanned_files_json, fw, indent=4)
                    fw.close()

                plyer.notification.notify(
                        title = "RIPHelper | Files scanned",
                        message = f'All files inside mentioned folders have been successfully scanned and {traces} new files have been found.',
                        timeout = 10,
                        app_name = "RIPHelper"
                )

            else:
            
                output += f"{date} Scan finished, no traces have been found. \n"

            with open(f"./logs/logs.txt", "a") as f:
                f.write(output)
                f.close()
                
            time.sleep(self.frequency)

if __name__ == "__main__":

    time.sleep(3)
    scanner = Scanner()
    scanner.scan()
    scanner.add_to_startup()


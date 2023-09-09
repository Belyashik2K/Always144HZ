import win32api
import time
import shutil as sh
import os

class RefreshRateUpdater:

    def __init__(self):
        """
        Initialize the script with the device name, target refresh rate, delay and script file name
        """
        self.device_name = win32api.EnumDisplayDevices().DeviceName
        self.target_rate = 144
        self.delay = 10
        self.script_file_name = __file__.split("\\")[-1]
        
    def add_to_startup(self):
        """
        Add the script to the startup folder if it is not there
        """
        if not os.path.exists(os.path.join("C:", os.environ["HOMEPATH"], "AppData", "Roaming", "Microsoft", "Windows", "Start Menu", "Programs", "Startup", self.script_file_name)):
            src = os.path.join(os.path.dirname(os.path.realpath(__file__)), self.script_file_name)
            dst = os.path.join("C:", os.environ["HOMEPATH"], "AppData", "Roaming", "Microsoft", "Windows", "Start Menu", "Programs", "Startup")
            sh.copy(src, dst)

    def get_rate(self):
        """
        Get the current refresh rate
        """
        settings = win32api.EnumDisplaySettings(self.device_name, -1)
        return settings.DisplayFrequency

    def polling_rate(self):
        """
        Polling the refresh rate with 10 second delay 
        and change it to 144Hz if it is not
        """
        while True:
            if self.get_rate() != self.target_rate:
                settings = win32api.EnumDisplaySettings(self.device_name, -1)
                settings.DisplayFrequency = self.target_rate
                win32api.ChangeDisplaySettings(settings, 0)
            time.sleep(self.delay)

def main():
    updater = RefreshRateUpdater()
    updater.add_to_startup()
    updater.polling_rate()

if __name__ == "__main__":
    main()
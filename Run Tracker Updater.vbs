Set objShell = CreateObject("WScript.Shell")
objShell.CurrentDirectory = "C:\AeroTracker"
objShell.Run "cmd /k ""C:\Users\EmilyWalton\AppData\Local\Microsoft\WindowsApps\python3.12.exe"" update_tracker.py", 1, False

"""Quick script to list all Outlook folders so we can find the right path."""
import pythoncom
pythoncom.CoInitialize()
import win32com.client

outlook = win32com.client.Dispatch("Outlook.Application")
ns = outlook.GetNamespace("MAPI")

def walk(folder, prefix=""):
    try:
        name = folder.Name
        print(f"{prefix}{name}")
        for i in range(1, folder.Folders.Count + 1):
            try:
                walk(folder.Folders.Item(i), prefix + "  ")
            except Exception:
                pass
    except Exception:
        pass

for i in range(1, ns.Folders.Count + 1):
    try:
        walk(ns.Folders.Item(i))
    except Exception:
        pass

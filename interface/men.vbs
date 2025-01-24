Set objShell = CreateObject("WScript.Shell")
objShell.Run "cmd /k cd /d C:\Migracion\interface && python menu.py", 1, True

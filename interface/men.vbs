Set objShell = CreateObject("WScript.Shell")
objShell.Run "cmd /c cd /d C:\Migracion\interface && python menu.py", 0, True

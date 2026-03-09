Set WshShell = CreateObject("WScript.Shell")

' Start SAP Logon
WshShell.Run """C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe""", 1, False
WScript.Sleep 5000

' Start Web Server
WshShell.Run "cmd /k cd /d C:\Project\Auto_SAP\Auto_SAP\sap_auto && C:\Project\Auto_SAP\Auto_SAP\venv\Scripts\python.exe manage.py runserver 0.0.0.0:8010", 7, False
WScript.Sleep 3000

' Start Scanner
WshShell.Run "cmd /k cd /d C:\Project\Auto_SAP\Auto_SAP\sap_auto && C:\Project\Auto_SAP\Auto_SAP\venv\Scripts\python.exe manage.py scan_tasks", 7, False
WScript.Sleep 3000

' Start Monitor (giám sát và tự restart nếu bị tắt)
WshShell.Run "wscript.exe ""C:\Project\Auto_SAP\Auto_SAP\monitor_all.vbs""", 0, False

MsgBox "SAP Auto Tasks started!", vbInformation, "SAP Auto"
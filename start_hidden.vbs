Set WshShell = CreateObject("WScript.Shell")

' Start SAP GUI
'WshShell.Run """C:\Program Files (x86)\SAP\FrontEnd\SAPgui\sapshcut.exe"" -system=V2Q -client=250 -user=ITS-2015030 -pw=Trung999", 1, False
'WScript.Sleep 20000

' Start Web Server (hidden)
WshShell.Run "cmd /c cd /d C:\Project\Auto_SAP\Auto_SAP\sap_auto && C:\Project\Auto_SAP\Auto_SAP\venv\Scripts\python.exe manage.py runserver 0.0.0.0:8010", 0, False
WScript.Sleep 5000

' Start Scanner (hidden)
WshShell.Run "cmd /c cd /d C:\Project\Auto_SAP\Auto_SAP\sap_auto && C:\Project\Auto_SAP\Auto_SAP\venv\Scripts\python.exe manage.py scan_tasks", 0, False

MsgBox "SAP Auto Tasks started!", vbInformation, "SAP Auto"
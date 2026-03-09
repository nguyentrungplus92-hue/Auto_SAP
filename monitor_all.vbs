Set WshShell = CreateObject("WScript.Shell")

Do While True
    ' === Kiểm tra SAP Logon ===
    bSapFound = False
    Set objWMI = GetObject("winmgmts:\\.\root\cimv2")
    Set colProcesses = objWMI.ExecQuery("SELECT * FROM Win32_Process WHERE Name = 'saplogon.exe' OR Name = 'sapgui.exe'")
    
    For Each objProcess In colProcesses
        If objProcess.SessionId > 0 Then
            bSapFound = True
            Exit For
        End If
    Next
    
    If Not bSapFound Then
        WshShell.Run """C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe""", 1, False
        WScript.Sleep 5000
    End If
    
    ' === Kiểm tra Web Server và Scanner ===
    nPythonCount = 0
    Set colPython = objWMI.ExecQuery("SELECT * FROM Win32_Process WHERE Name = 'python.exe'")
    
    For Each objProcess In colPython
        If objProcess.SessionId > 0 Then
            nPythonCount = nPythonCount + 1
        End If
    Next
    
    ' Cần ít nhất 2 python (Web + Scanner)
    If nPythonCount < 2 Then
        ' Restart cả 2
        WshShell.Run "cmd /k cd /d C:\Project\Auto_SAP\Auto_SAP\sap_auto && C:\Project\Auto_SAP\Auto_SAP\venv\Scripts\python.exe manage.py runserver 0.0.0.0:8010", 1, False
        WScript.Sleep 3000
        WshShell.Run "cmd /k cd /d C:\Project\Auto_SAP\Auto_SAP\sap_auto && C:\Project\Auto_SAP\Auto_SAP\venv\Scripts\python.exe manage.py scan_tasks", 1, False
        WScript.Sleep 5000
    End If
    
    ' Kiểm tra mỗi 30 giây
    WScript.Sleep 30000
Loop
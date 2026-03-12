Set WshShell = CreateObject("WScript.Shell")
Set objWMI = GetObject("winmgmts:\\.\root\cimv2")

Do While True
    ' === Kiểm tra SAP Logon (chỉ SessionId > 0) ===
    bSapFound = False
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
    
    ' === Kiểm tra Web Server (runserver 8010, chỉ SessionId > 0) ===
    bWebFound = False
    Set colPython = objWMI.ExecQuery("SELECT * FROM Win32_Process WHERE Name = 'python.exe'")
    
    For Each objProcess In colPython
        If objProcess.SessionId > 0 Then
            strCmd = ""
            On Error Resume Next
            strCmd = objProcess.CommandLine
            On Error Goto 0
            
            If InStr(strCmd, "runserver") > 0 And InStr(strCmd, "8010") > 0 Then
                bWebFound = True
                Exit For
            End If
        End If
    Next
    
    If Not bWebFound Then
        WshShell.Run "cmd /k cd /d C:\Project\Auto_SAP\Auto_SAP\sap_auto && C:\Project\Auto_SAP\Auto_SAP\venv\Scripts\python.exe manage.py runserver 0.0.0.0:8010", 7, False
        WScript.Sleep 5000
    End If
    
    ' === Kiểm tra Scanner (scan_tasks, chỉ SessionId > 0) ===
    bScannerFound = False
    Set colPython2 = objWMI.ExecQuery("SELECT * FROM Win32_Process WHERE Name = 'python.exe'")
    
    For Each objProcess In colPython2
        If objProcess.SessionId > 0 Then
            strCmd = ""
            On Error Resume Next
            strCmd = objProcess.CommandLine
            On Error Goto 0
            
            If InStr(strCmd, "scan_tasks") > 0 Then
                bScannerFound = True
                Exit For
            End If
        End If
    Next
    
    If Not bScannerFound Then
        WshShell.Run "cmd /k cd /d C:\Project\Auto_SAP\Auto_SAP\sap_auto && C:\Project\Auto_SAP\Auto_SAP\venv\Scripts\python.exe manage.py scan_tasks", 7, False
        WScript.Sleep 5000
    End If
    
    ' Kiểm tra mỗi 30 giây
    WScript.Sleep 30000
Loop
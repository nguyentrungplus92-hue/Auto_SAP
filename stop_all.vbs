Set WshShell = CreateObject("WScript.Shell")

' Kill all python processes
WshShell.Run "taskkill /F /IM python.exe", 0, True

MsgBox "SAP Auto Tasks stopped!", vbInformation, "SAP Auto"
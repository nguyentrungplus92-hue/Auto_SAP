Set WshShell = CreateObject("WScript.Shell")

' Kill monitor trước (có thể có nhiều instance)
WshShell.Run "taskkill /F /IM wscript.exe", 0, True
WScript.Sleep 1000

' Kill lần nữa để chắc chắn
WshShell.Run "taskkill /F /IM wscript.exe", 0, True
WScript.Sleep 1000

' Kill python
WshShell.Run "taskkill /F /IM python.exe", 0, True
WScript.Sleep 1000

MsgBox "SAP Auto Tasks stopped!", vbInformation, "SAP Auto"
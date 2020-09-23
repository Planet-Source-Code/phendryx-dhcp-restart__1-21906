Attribute VB_Name = "basIni"
Public iniAppName As String                 'Ini file application name
Public iniFileName As String                'Ini file filename

Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpString As Any, ByVal lpFileName As String) As Long

Public Function IniRead(iniKeyName As String)

Dim MyStr As String

On Local Error GoTo ErrorHandler

MyStr = String(255, Chr(0))

GetPrivateProfileString iniAppName, iniKeyName, "", MyStr, Len(MyStr) - 1, iniFileName
IniRead = Left(MyStr, InStr(1, MyStr, Chr(0)) - 1)
Exit Function

ErrorHandler:
Select Case Err.Number
    Case Else
        MsgBox "Error " & Err.Number & " (" & Err.Description & ") occured in basIni:General:IniRead"
End Select

End Function

Public Sub IniWrite(iniKeyName As String, iniKeyValue As String)

On Local Error GoTo ErrorHandler

WritePrivateProfileString iniAppName, iniKeyName, iniKeyValue, iniFileName
Exit Sub

ErrorHandler:
Select Case Err.Number
    Case Else
        MsgBox "Error " & Err.Number & " (" & Err.Description & ") occured in basIni:General:IniWrite"
End Select

End Sub

Public Sub IniSettings(AppName As String, FileName As String)

On Local Error GoTo ErrorHandler

iniAppName = AppName
iniFileName = FileName
Exit Sub

ErrorHandler:
Select Case Err.Number
    Case Else
        MsgBox "Error " & Err.Number & " (" & Err.Description & ") occured in basIni:General:IniSettings"
End Select

End Sub

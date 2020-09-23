Attribute VB_Name = "basShellAndWait"
Sub ShellAndWait(lstrShellString As String, lbolVisible As Boolean)

Dim objScript

Set objScript = CreateObject("WScript.Shell")

If lbolVisible = True Then
    ShellApp = objScript.Run(lstrShellString, 1, True)
Else
    ShellApp = objScript.Run(lstrShellString, 0, True)
End If

End Sub

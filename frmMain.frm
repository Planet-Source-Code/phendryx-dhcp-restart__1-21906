VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DHCP Restart"
   ClientHeight    =   2925
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   4680
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2925
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.Frame frameLog 
      Caption         =   "Log"
      Height          =   1215
      Left            =   120
      TabIndex        =   7
      Top             =   1680
      Width           =   4455
      Begin VB.ListBox lstLog 
         Height          =   840
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   4215
      End
   End
   Begin VB.Frame frameStats 
      Caption         =   "Stats"
      Height          =   1695
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   4455
      Begin MSWinsockLib.Winsock wskBrowser 
         Left            =   3960
         Top             =   1080
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin VB.Timer tmrCurrentTime 
         Interval        =   1000
         Left            =   3960
         Top             =   600
      End
      Begin VB.Timer tmrOneSecond 
         Interval        =   1000
         Left            =   3960
         Top             =   120
      End
      Begin VB.Label lblCurrentTime 
         Caption         =   "1/1/2001 12:00:00 am"
         Height          =   255
         Left            =   1200
         TabIndex        =   10
         Top             =   1320
         Width           =   3135
      End
      Begin VB.Label lblCurrentTimeCaption 
         Alignment       =   1  'Right Justify
         Caption         =   "Current Time:"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   1320
         Width           =   975
      End
      Begin VB.Label lblLastCheck 
         Caption         =   "1/1/2001 12:00:00 am"
         Height          =   255
         Left            =   1200
         TabIndex        =   6
         Top             =   600
         Width           =   3135
      End
      Begin VB.Label lblStatus 
         Caption         =   "Idle..."
         Height          =   255
         Left            =   1200
         TabIndex        =   5
         Top             =   240
         Width           =   3135
      End
      Begin VB.Label lblNextCheck 
         Caption         =   "1/2/2001 12:00:00 am"
         Height          =   255
         Left            =   1200
         TabIndex        =   4
         Top             =   960
         Width           =   3135
      End
      Begin VB.Label lblNextCheckCaption 
         Alignment       =   1  'Right Justify
         Caption         =   "Next Check:"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   960
         Width           =   975
      End
      Begin VB.Label lblLastCheckCaption 
         Alignment       =   1  'Right Justify
         Caption         =   "Last Check:"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   600
         Width           =   975
      End
      Begin VB.Label lblStatusCaption 
         Caption         =   "Status:"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Menu mPopupSys 
      Caption         =   "mPopupSys"
      Visible         =   0   'False
      Begin VB.Menu mnumPopupSysShowForm 
         Caption         =   "Show Form"
      End
      Begin VB.Menu mnumPopupSysExit 
         Caption         =   "&End"
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileSettings 
         Caption         =   "Settings"
      End
      Begin VB.Menu mnuFileSeperator1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuViewHideForm 
         Caption         =   "Hide"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

Me.lblStatus.Caption = "Starting..."

'the form must be fully visible before calling Shell_NotifyIcon
Me.Show
Me.Refresh
With nid
    .cbSize = Len(nid)
    .hwnd = Me.hwnd
    .uId = vbNull
    .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
    .uCallBackMessage = WM_MOUSEMOVE
    .hIcon = Me.Icon
    .szTip = "Your ToolTip" & vbNullChar
End With
Shell_NotifyIcon NIM_ADD, nid

Me.WindowState = vbMinimized

gstrAppName = "DHCP Restart"

IniSettings gstrAppName, App.Path & "\" & gstrAppName & ".ini"

Me.Caption = gstrAppName & " v" & App.Major & "." & App.Minor & "." & App.Revision

frmMain.lstLog.AddItem "Started: " & Now
frmMain.lblNextCheck.Caption = DateAdd("s", 5, Now)

Me.lblStatus.Caption = "Checking Settings..."
If IniRead("CheckInterval") = "" Or IniRead("URLToLoad") = "" Then
    ShowSettings
End If

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

      'this procedure receives the callbacks from the System Tray icon.
      Dim Result As Long
      Dim msg As Long
       'the value of X will vary depending upon the scalemode setting
       If Me.ScaleMode = vbPixels Then
        msg = X
       Else
        msg = X / Screen.TwipsPerPixelX
       End If
       Select Case msg
        'case WM_LBUTTONUP        '514 restore form window
        ' Me.WindowState = vbNormal
        ' Result = SetForegroundWindow(Me.hWnd)
        ' Me.Show
        Case WM_LBUTTONDBLCLK    '515 restore form window
         Me.WindowState = vbNormal
         Result = SetForegroundWindow(Me.hwnd)
         Me.Show
        Case WM_RBUTTONUP        '517 display popup menu
         Result = SetForegroundWindow(Me.hwnd)
         Me.PopupMenu Me.mPopupSys
       End Select

End Sub

Private Sub Form_Resize()

If Me.WindowState = vbMinimized Then
    Me.Visible = False
End If

End Sub

Private Sub Form_Unload(Cancel As Integer)

'this removes the icon from the system tray
Shell_NotifyIcon NIM_DELETE, nid

End Sub

Private Sub mnuFileExit_Click()

Unload Me
End

End Sub

Private Sub mnuFileSettings_Click()

frmSettings.Show vbModal

End Sub

Private Sub mnuHelpAbout_Click()

frmAbout.Show vbModal

End Sub

Private Sub mnumPopupSysExit_Click()

Unload Me
End

End Sub

Private Sub mnumPopupSysShowForm_Click()

Me.Visible = True

End Sub

Private Sub mnuViewHideForm_Click()

Me.Visible = False

End Sub

Private Sub tmrCurrentTime_Timer()

Me.lblCurrentTime.Caption = Now

End Sub

Private Sub tmrOneSecond_Timer()

Me.tmrOneSecond.Enabled = False

If Me.lblNextCheck.Caption = Now Then
    CheckDHCPConnection
Else
    Me.lblStatus.Caption = "Waiting..."
    Me.tmrOneSecond.Enabled = True
End If

End Sub

Public Sub CheckDHCPConnection()

Dim lstrShellString As String
Dim lbolReplyFrom As Boolean
Dim lintInFile As Integer
Dim lstrTemp As String
Dim lintOutFile As Integer

Me.lblStatus.Caption = "Pinging..."
lintOutFile = FreeFile
Open App.Path & "\DHCPPing.bat" For Output As #lintOutFile
    Print #lintOutFile, lstrShellString
Close #lintOutFile

lstrShellString = "ping -n 1 " & Replace(IniRead("URLToLoad"), "http://", "") & " > """ & App.Path & "\ping.tmp"""

lintOutFile = FreeFile
Open App.Path & "\DHCPPing.bat" For Output As #lintOutFile
    Print #lintOutFile, lstrShellString
Close #lintOutFile

ShellAndWait """" & App.Path & "\DHCPPing.bat""", False

lintInFile = FreeFile
Open App.Path & "\ping.tmp" For Input As #lintInFile
    Do Until EOF(lintInFile)
        If InStr(1, LCase(lstrTemp), "reply from") > 0 Then
            lbolReplyFrom = True
            Exit Do
        End If
        
        Line Input #lintInFile, lstrTemp
    Loop
Close #lintInFile

Kill App.Path & "\DHCPPing.bat"
Kill App.Path & "\ping.tmp"

If lbolReplyFrom = True Then
    Me.lblStatus.Caption = "Ping successful..."
    PrintToLog "Check: " & Now & " - Succeeded"
    Me.wskBrowser.Close
    
    Select Case IniRead("CheckInterval")
        Case "1 Minute"
            Me.lblLastCheck.Caption = Now
            Me.lblNextCheck.Caption = DateAdd("n", 1, Now)
        Case "15 Minutes"
            Me.lblLastCheck.Caption = Now
            Me.lblNextCheck.Caption = DateAdd("n", 15, Now)
        Case "30 Minutes"
            Me.lblLastCheck.Caption = Now
            Me.lblNextCheck.Caption = DateAdd("n", 30, Now)
        Case "45 Minutes"
            Me.lblLastCheck.Caption = Now
            Me.lblNextCheck.Caption = DateAdd("n", 45, Now)
        Case "1 Hour"
            Me.lblLastCheck.Caption = Now
            Me.lblNextCheck.Caption = DateAdd("h", 1, Now)
        Case "1 Day"
            Me.lblLastCheck.Caption = Now
            Me.lblNextCheck.Caption = DateAdd("d", 1, Now)
    End Select
    Me.tmrOneSecond.Enabled = True
Else
    Me.lblStatus.Caption = "Ping failure..."
    PrintToLog "Check: " & Now & " - Failed"
    Me.wskBrowser.Close
    
    Me.lblStatus.Caption = "Releasing DHCP Info..."
    PrintToLog "DHCP: " & Now & " - Release"
    lstrShellString = "ipconfig /release"
    ShellAndWait lstrShellString, False
    
    Me.lblStatus.Caption = "Renewing DHCP Info..."
    PrintToLog "DHCP: " & Now & " - Renew"
    lstrShellString = "ipconfig /renew"
    ShellAndWait lstrShellString, False
    
    Me.lblStatus.Caption = "Waiting..."
    Me.lblNextCheck.Caption = DateAdd("s", 5, Now)
    Me.tmrOneSecond.Enabled = True
End If


End Sub

Public Sub PrintToLog(lstrLogText As String)

Me.lstLog.AddItem lstrLogText
If Me.lstLog.ListCount = 51 Then
    Me.lstLog.RemoveItem 0
End If
Me.lstLog.Selected(Me.lstLog.ListCount - 1) = True

End Sub

Public Sub ShowSettings()

frmSettings.Show vbModal
MsgBox "Please restart " & gstrAppName & ".", vbOKOnly, gstrAppName

'this removes the icon from the system tray
Shell_NotifyIcon NIM_DELETE, nid

End

End Sub

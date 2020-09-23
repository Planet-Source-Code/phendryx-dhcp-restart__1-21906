VERSION 5.00
Begin VB.Form frmSettings 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Settings"
   ClientHeight    =   1665
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   Icon            =   "frmSettings.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1665
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2400
      TabIndex        =   3
      Top             =   1200
      Width           =   1095
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   1080
      TabIndex        =   2
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Frame frameSettings 
      Caption         =   "Settings"
      Height          =   1095
      Left            =   120
      TabIndex        =   4
      Top             =   0
      Width           =   4455
      Begin VB.ComboBox cmbCheckInterval 
         Height          =   315
         ItemData        =   "frmSettings.frx":0442
         Left            =   1320
         List            =   "frmSettings.frx":0444
         TabIndex        =   1
         Top             =   720
         Width           =   3015
      End
      Begin VB.TextBox txtURLToLoad 
         Height          =   285
         Left            =   1320
         TabIndex        =   0
         Top             =   240
         Width           =   3015
      End
      Begin VB.Label lblCheckIntervalCaption 
         Alignment       =   1  'Right Justify
         Caption         =   "Check Interval:"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label lblURLToLoadCaption 
         Alignment       =   1  'Right Justify
         Caption         =   "URL To Load:"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frmSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()


End Sub

Private Sub cmdSave_Click()

IniWrite "URLToLoad", Me.txtURLToLoad.Text
IniWrite "CheckInterval", Me.cmbCheckInterval.Text
Unload Me

End Sub

Private Sub Command2_Click()

Unload Me

End Sub

Private Sub Form_Load()

Me.Caption = gstrAppName & " - " & Me.Caption

Dim lintI As Integer

With Me.cmbCheckInterval
    .AddItem "1 Minute"
    .AddItem "15 Minutes"
    .AddItem "30 Minutes"
    .AddItem "45 Minutes"
    .AddItem "1 Hour"
    .AddItem "1 Day"
End With

Me.txtURLToLoad.Text = IniRead("URLToLoad")
For lintI = 0 To Me.cmbCheckInterval.ListCount - 1
    If Me.cmbCheckInterval.List(lintI) = IniRead("CheckInterval") Then
        Me.cmbCheckInterval.ListIndex = lintI
    End If
Next lintI

End Sub

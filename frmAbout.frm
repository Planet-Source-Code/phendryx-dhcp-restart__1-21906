VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About"
   ClientHeight    =   1395
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "frmAbout.frx":0442
   ScaleHeight     =   1395
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frameAbout 
      Height          =   1335
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   4455
      Begin VB.Label lblDHCPRestart 
         Alignment       =   2  'Center
         Caption         =   "DHCP Restart"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   960
         Width           =   4215
      End
      Begin VB.Label lblLink 
         Alignment       =   2  'Center
         Caption         =   "http://www.tekker.com"
         Height          =   255
         Left            =   120
         MouseIcon       =   "frmAbout.frx":074C
         MousePointer    =   99  'Custom
         TabIndex        =   2
         Top             =   600
         Width           =   4215
      End
      Begin VB.Label lblTekkerSoft 
         Alignment       =   2  'Center
         Caption         =   "TekkerSoft"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   4215
      End
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

Me.lblDHCPRestart.Caption = gstrAppName & " v" & App.Major & "." & App.Minor & "." & App.Revision

Me.Caption = gstrAppName & " - " & Me.Caption

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Me.lblLink.ForeColor = vbBlack

End Sub

Private Sub frameAbout_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Me.lblLink.ForeColor = vbBlack

End Sub

Private Sub lblLink_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

Shell "start http://www.tekker.com"

End Sub

Private Sub lblLink_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Me.lblLink.ForeColor = vbBlue

End Sub


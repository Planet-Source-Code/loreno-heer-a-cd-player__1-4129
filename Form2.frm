VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form ShapedForm 
   BorderStyle     =   0  'Kein
   Caption         =   "Player"
   ClientHeight    =   2955
   ClientLeft      =   120
   ClientTop       =   120
   ClientWidth     =   4410
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   Picture         =   "Form2.frx":0000
   ScaleHeight     =   2955
   ScaleWidth      =   4410
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.OptionButton Option3 
      Caption         =   "Option3"
      Height          =   195
      Left            =   3480
      TabIndex        =   13
      Top             =   1920
      Width           =   255
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Option2"
      Height          =   195
      Left            =   2880
      TabIndex        =   11
      Top             =   2160
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Height          =   225
      Left            =   2040
      TabIndex        =   8
      Text            =   "1"
      Top             =   1800
      Width           =   735
   End
   Begin MP3Player.AXMarquee AXMarquee2 
      Height          =   690
      Left            =   360
      TabIndex        =   4
      Top             =   840
      Width           =   3975
      _extentx        =   7011
      _extenty        =   1217
      text            =   "Programmiert von Loreno Heer     Borgs best Player"
   End
   Begin MP3Player.AXMarquee AXMarquee1 
      Height          =   690
      Left            =   360
      TabIndex        =   3
      Top             =   240
      Width           =   3975
      _extentx        =   7011
      _extenty        =   1217
      text            =   "Borgs best Player Programmiert von Loreno Heer"
   End
   Begin MSComCtl2.FlatScrollBar FlatScrollBar1 
      Height          =   855
      Left            =   360
      TabIndex        =   6
      Top             =   1560
      Width           =   135
      _ExtentX        =   238
      _ExtentY        =   1508
      _Version        =   393216
      Appearance      =   2
      Orientation     =   8323072
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Option1"
      Height          =   195
      Left            =   2040
      TabIndex        =   15
      Top             =   2160
      Width           =   255
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Alle"
      Height          =   195
      Left            =   1680
      TabIndex        =   10
      Top             =   2160
      Width           =   255
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Shuffle"
      Height          =   195
      Left            =   2880
      TabIndex        =   14
      Top             =   1920
      Width           =   495
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Einzeln"
      Height          =   195
      Left            =   2280
      TabIndex        =   12
      Top             =   2160
      Width           =   510
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Track:"
      Height          =   195
      Left            =   1560
      TabIndex        =   9
      Top             =   1800
      Width           =   465
   End
   Begin VB.Line Line1 
      X1              =   1560
      X2              =   4080
      Y1              =   1800
      Y2              =   1800
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CD-Player"
      Height          =   195
      Left            =   1560
      TabIndex        =   7
      Top             =   1560
      Width           =   705
   End
   Begin VB.Image Image4 
      Height          =   210
      Left            =   1200
      Picture         =   "Form2.frx":2981A
      Top             =   2040
      Width           =   315
   End
   Begin VB.Image Image3 
      Height          =   210
      Left            =   720
      Picture         =   "Form2.frx":29BDC
      Top             =   2040
      Width           =   315
   End
   Begin VB.Image Image2 
      Height          =   210
      Left            =   1200
      Picture         =   "Form2.frx":29F9E
      Top             =   1680
      Width           =   315
   End
   Begin VB.Image Image1 
      Height          =   210
      Left            =   720
      Picture         =   "Form2.frx":2A360
      Top             =   1680
      Width           =   315
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   2520
      Width           =   4095
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   3720
      TabIndex        =   2
      Top             =   0
      Width           =   135
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   3960
      TabIndex        =   1
      Top             =   0
      Width           =   255
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4215
   End
End
Attribute VB_Name = "ShapedForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Form_Load()
Me.AXMarquee1.Scrolling = True
MenuForm.Hide
End Sub

Private Sub Image1_Click()
If Me.Option1.Value = True Then
Module2.playCdFull
ElseIf Me.Option2.Value = True Then
Module2.playCdTrack (Me.Text1.Text)
ElseIf Me.Option3.Value = True Then
Module2.playCdShuffle
End If
End Sub

Private Sub Image2_Click()
If MenuForm.Text1.Text = "1" Then
    Module2.DeMute
    MenuForm.Text1.Text = 0
Else
    Module2.mute
    MenuForm.Text1.Text = 1
End If
End Sub

Private Sub Image3_Click()
Module2.stopCd
End Sub

Private Sub Image4_Click()
If Me.AXMarquee2.Text = "CD Door open" Then
Module2.doorClose
Me.AXMarquee2.Text = ""
Me.AXMarquee2.Scrolling = False
Else
Me.AXMarquee2.Text = "CD Door open"
Me.AXMarquee2.Scrolling = True
Module2.doorOpen
End If
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
ReleaseCapture
SendMessage Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End Sub

Private Sub Label2_Click()
Unload MenuForm
Unload Me
End Sub

Private Sub Label3_Click()
ShapedForm.WindowState = 1
End Sub


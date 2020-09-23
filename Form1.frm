VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form MenuForm 
   BorderStyle     =   4  'Festes Werkzeugfenster
   Caption         =   "Player"
   ClientHeight    =   2100
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3675
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   2100
   ScaleWidth      =   3675
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   6120
      TabIndex        =   4
      Text            =   "Text5"
      Top             =   1200
      Width           =   1335
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   6120
      TabIndex        =   3
      Text            =   "Text4"
      Top             =   960
      Width           =   1335
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   6120
      TabIndex        =   2
      Text            =   "Text3"
      Top             =   720
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   6120
      TabIndex        =   1
      Text            =   "Text2"
      Top             =   480
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   6120
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   240
      Width           =   1335
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   120
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.Menu mOptions 
      Caption         =   "&Options"
      Begin VB.Menu mChangeBackgroundPicture 
         Caption         =   "Change Background &Picture"
      End
      Begin VB.Menu mInstructions 
         Caption         =   "&Instructions"
      End
      Begin VB.Menu mSeparator 
         Caption         =   "-"
      End
      Begin VB.Menu mExit 
         Caption         =   "&Exit"
      End
   End
End
Attribute VB_Name = "MenuForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit


Private hRgn As Long

Private Const OFN_FILEMUSTEXIST = &H1000
Private Const OFN_HIDEREADONLY = &H4
Private Const OFN_LONGNAMES = &H200000
Private Const OFN_NONETWORKBUTTON = &H20000
Private Const OFN_PATHMUSTEXIST = &H800
Private Const CC_FULLOPEN = &H2
Private Const CC_SOLIDCOLOR = &H80
Private Const CC_RGBINIT = &H1
Private Const CC_ANYCOLOR = &H100

Private Sub Form_Load()
    CommonDialog1.Color = vbWhite
    SetRegion
    ShapedForm.Show
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If hRgn Then DeleteObject hRgn
    Unload ShapedForm
End Sub

Private Sub mChangeBackgroundPicture_Click()
    On Error Resume Next
    Err.Clear
    With CommonDialog1
        .DialogTitle = "Please Select a Picture"
        .Flags = OFN_FILEMUSTEXIST + OFN_HIDEREADONLY + OFN_LONGNAMES + OFN_NONETWORKBUTTON + OFN_PATHMUSTEXIST
        .Filter = "All Picture Files|*.bmp;*.dib;*.gif;*.jpg;*.wmf;*.emf;*.ico;*.cur|Bitmaps (*.bmp;*.dib)|*.bmp;*.dib|GIF Images (*.gif)|*.gif|JPEG Images (*.jpg)|*.jpg|Metafiles (*.wmf;*.emf)|*.wmf;*.emf|Icons (*.ico;*.cur)|*.ico;*.cur|All Files (*.*)|*.*"
        .ShowOpen
        If Err.Number = 32755 Then Exit Sub
        .Flags = CC_FULLOPEN + CC_SOLIDCOLOR + CC_RGBINIT + CC_ANYCOLOR
        .ShowColor
        If Err.Number = 32755 Then Exit Sub
        On Error GoTo erro
        ShapedForm.Visible = False
        DoEvents
        ShapedForm.Picture = LoadPicture(.FileName)
        ShapedForm.Width = ShapedForm.Picture.Width
        ShapedForm.Height = ShapedForm.Picture.Height
        SetRegion
    End With
erro:
    If Err.Number <> 0 Then MsgBox "Error Number " & Err.Number & " : " & Err.Description, vbApplicationModal + vbCritical
    ShapedForm.Visible = True
End Sub

Private Sub mExit_Click()
    Unload Me
End Sub

Private Sub SetRegion()
    If hRgn Then DeleteObject hRgn
    hRgn = GetBitmapRegion(ShapedForm.Picture, CommonDialog1.Color)
    SetWindowRgn ShapedForm.hwnd, hRgn, True
End Sub

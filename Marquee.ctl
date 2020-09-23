VERSION 5.00
Begin VB.UserControl AXMarquee 
   Appearance      =   0  '2D
   AutoRedraw      =   -1  'True
   BackColor       =   &H8000000D&
   ClientHeight    =   2730
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4605
   PropertyPages   =   "Marquee.ctx":0000
   ScaleHeight     =   182
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   307
   ToolboxBitmap   =   "Marquee.ctx":0011
   Begin VB.PictureBox picBlankCol 
      Appearance      =   0  '2D
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'Kein
      ForeColor       =   &H80000008&
      Height          =   690
      Left            =   720
      Picture         =   "Marquee.ctx":010B
      ScaleHeight     =   46
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   5
      TabIndex        =   2
      Top             =   600
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.PictureBox picCaps 
      Appearance      =   0  '2D
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'Kein
      ForeColor       =   &H80000008&
      Height          =   540
      Left            =   -2148
      Picture         =   "Marquee.ctx":06BD
      ScaleHeight     =   35.752
      ScaleMode       =   0  'Benutzerdefiniert
      ScaleWidth      =   889.6
      TabIndex        =   1
      Top             =   2130
      Width           =   13350
   End
   Begin VB.PictureBox picMsg 
      Appearance      =   0  '2D
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'Kein
      ForeColor       =   &H80000008&
      Height          =   540
      Left            =   0
      ScaleHeight     =   36
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   78
      TabIndex        =   0
      Top             =   1485
      Width           =   1170
   End
   Begin VB.Timer tAni 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   204
      Top             =   156
   End
End
Attribute VB_Name = "AXMarquee"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "ActiveX Marquee Control"
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
Option Explicit

Enum ScrollModeValue
  R_to_L = 0
  L_to_R = 1
End Enum

'Vars for tracking BMP size and position
Private lBMPWidth   As Long     'Total width of the Message Bitmap to be drawn on the background
Private bRestart    As Boolean
Private lCtlWidth   As Long
Const SRC_Y = 0
Const CTL_HEIGHT = 683
Const m_def_ScrollMode = R_to_L
Const m_def_Text = "ActiveX Marquee"
Const m_def_Scrolling = False
Dim m_ScrollMode As ScrollModeValue
Dim m_Text As String
Dim m_Scrolling As Boolean

Private Sub tAni_Timer()
  Static lX           As Long
  Static lX2          As Long
  Static lSrcOffset   As Long
  Static lSrcWidth    As Long

  If bRestart Then
    If m_ScrollMode = R_to_L Then
      lX = lCtlWidth - BULB_WIDTH
      lSrcOffset = 0
      lSrcWidth = BULB_WIDTH
    Else
    
      lX = BULB_WIDTH
      lSrcOffset = BULB_WIDTH
      lSrcWidth = BULB_WIDTH
    End If

    bRestart = False
  End If
  
  If m_ScrollMode = R_to_L Then
    If lX > 0 Then
      lX2 = lX
      If lCtlWidth - lX <= lBMPWidth Then
        lSrcWidth = lCtlWidth - lX
      Else
        lSrcWidth = lBMPWidth
      End If
    Else
      lX2 = 0
      lSrcOffset = Abs(lX)
      lSrcWidth = lBMPWidth - lSrcOffset
    End If
  Else
    If lX < lCtlWidth Then
      If lX <= lBMPWidth Then
        lX2 = 0
        lSrcWidth = lX
        lSrcOffset = lBMPWidth - lX
      Else
        lX2 = lX2 + BULB_WIDTH
        lSrcWidth = lBMPWidth
        lSrcOffset = 0
      End If
    Else
      If lX > lBMPWidth Then
        lX2 = lX2 + BULB_WIDTH
        lSrcWidth = lBMPWidth
      Else
        lSrcOffset = lBMPWidth - lX
        lSrcWidth = lCtlWidth
      End If
    End If
  End If
  
  UserControl.PaintPicture picMsg.Picture, lX2, SRC_Y, , , _
                           lSrcOffset, , lSrcWidth, , _
                           vbSrcCopy
  
  If m_ScrollMode = R_to_L Then
    If lSrcOffset + BULB_WIDTH = lBMPWidth Then
      bRestart = True
    Else
      lX = lX - BULB_WIDTH
    End If
  Else
    If lX2 + BULB_WIDTH = lCtlWidth Then
      bRestart = True
    Else
      lX = lX + BULB_WIDTH
    End If
  End If
  
End Sub

Private Sub UserControl_Initialize()
  InitBMPStruct
End Sub
Private Sub UserControl_InitProperties()
  m_ScrollMode = m_def_ScrollMode
  m_Text = m_def_Text
  m_Scrolling = m_def_Scrolling
End Sub
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
  ScrollMode = PropBag.ReadProperty("ScrollMode", m_def_ScrollMode)
  Text = PropBag.ReadProperty("Text", m_def_Text)
  Scrolling = PropBag.ReadProperty("Scrolling", m_def_Scrolling)
End Sub
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
  Call PropBag.WriteProperty("ScrollMode", m_ScrollMode, m_def_ScrollMode)
  Call PropBag.WriteProperty("Text", m_Text, m_def_Text)
  Call PropBag.WriteProperty("Scrolling", m_Scrolling, m_def_Scrolling)
End Sub

Private Sub UserControl_Resize()
  UserControl.Height = CTL_HEIGHT
  lCtlWidth = UserControl.ScaleWidth - UserControl.ScaleWidth Mod 5
  DrawBackground
End Sub

Public Property Get Text() As String
Attribute Text.VB_Description = "Text string to display on the marquee"
Attribute Text.VB_ProcData.VB_Invoke_Property = ";Text"
  Text = m_Text
End Property

Public Property Let Text(ByVal New_Text As String)
  m_Text = New_Text
  PropertyChanged "Text"
  If m_Scrolling Then
    tAni.Enabled = False
    bRestart = True
    DrawBackground
    BuildTheBmp (m_Text)
    tAni.Enabled = True
  Else
    tAni.Enabled = False
    bRestart = False
  End If

End Property

Public Property Get Scrolling() As Boolean
Attribute Scrolling.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
Attribute Scrolling.VB_ProcData.VB_Invoke_Property = ";Behavior"
  Scrolling = m_Scrolling
End Property

Public Property Let Scrolling(ByVal bScrolling As Boolean)
  
  m_Scrolling = bScrolling
  
  PropertyChanged "Scrolling"
  
  If m_Scrolling Then
    DrawBackground
    BuildTheBmp (m_Text)
    tAni.Enabled = True
  Else
    tAni.Enabled = False
    bRestart = False
  End If
  
End Property

Public Property Get ScrollMode() As ScrollModeValue
  ScrollMode = m_ScrollMode
End Property

Public Property Let ScrollMode(ByVal New_ScrollMode As ScrollModeValue)
  m_ScrollMode = New_ScrollMode
  PropertyChanged "ScrollMode"
  If m_Scrolling Then
    tAni.Enabled = False
    bRestart = True
    DrawBackground
    BuildTheBmp (m_Text)
    tAni.Enabled = True
  Else
    tAni.Enabled = False
    bRestart = False
  End If

End Property

Private Sub DrawBackground()
  Dim lColX As Long
  
  With UserControl
        .AutoRedraw = True
    
    For lColX = 0 To .ScaleWidth Step 5
    
      .PaintPicture picBlankCol.Picture, lColX, 0, _
                    aCharSpace.Width, , _
                    aCharSpace.Left, 0, _
                    aCharSpace.Width
    
    Next lColX
    
    
    .AutoRedraw = False
    
  End With
  
End Sub

Private Function BuildTheBmp(sText As String) As Long
  Dim lChar     As Long
  Dim lOffset   As Long
  Dim lCharVal  As Long
  Dim lCounter  As Long
  Dim lMsgLength As Long
  
  
  sText = UCase$(sText)
  lMsgLength = Len(sText)
  
  With picMsg
  
      .AutoRedraw = True
    
      For lChar = 1 To lMsgLength
      lCharVal = Asc(Mid$(sText, lChar, 1))
      If lCharVal = 32 Then
        For lCounter = 1 To 4
          lOffset = lOffset + aCharSpace.Width
        Next lCounter
      
      ElseIf lCharVal >= 65 And lCharVal <= 90 Then
        lOffset = lOffset + aChars(lCharVal).Width
      End If
      
    Next lChar
    
    
    .Width = lOffset + aCharSpace.Width
    
    lOffset = 0
    
    For lChar = 1 To lMsgLength
      
    
      lCharVal = Asc(Mid$(sText, lChar, 1))
      
      If lCharVal = 32 Then
      
        For lCounter = 1 To 4
          .PaintPicture picCaps.Picture, lOffset, 0, _
                        aCharSpace.Width, , _
                        aCharSpace.Left, 0, _
                        aCharSpace.Width
          
          lOffset = lOffset + aCharSpace.Width

        Next lCounter
              
      ElseIf lCharVal >= 65 And lCharVal <= 90 Then
            
                
        .PaintPicture picCaps.Picture, lOffset, 0, _
                      aChars(lCharVal).Width, , _
                      aChars(lCharVal).Left, 0, _
                      aChars(lCharVal).Width
                      
        
        lOffset = lOffset + aChars(lCharVal).Width
      
      Else
        Debug.Print "Unsupported character entered - " & Mid$(sText, lChar, 1) & "ASCII = " & Asc(Mid$(sText, lChar, 1))
      
      End If
      
    Next lChar
    
    
    .PaintPicture picCaps.Picture, lOffset, 0, _
                  aCharSpace.Width, , _
                  aCharSpace.Left, 0, _
                  aCharSpace.Width
                  
    lOffset = lOffset + aCharSpace.Width
    
    
    .AutoRedraw = False
    
    .Picture = picMsg.Image
    
  End With
  
    lBMPWidth = lOffset
  
  BuildTheBmp = 0
  
  bRestart = True
End Function

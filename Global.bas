Attribute VB_Name = "GlobalStuff"
Option Explicit
Type CharBMP
  Left As Long
  Width As Long
End Type
Public aCharSpace As CharBMP
Public aChars(65 To 90) As CharBMP
Public Const BULB_WIDTH = 5
Public Const CHAR_HEIGHT = 36
Public Const CHAR_THIN = 30
Public Const CHAR_WIDE = 35

Public Sub InitBMPStruct()
  Dim x As Long
  aCharSpace.Left = 0
  aCharSpace.Width = 5
  aChars(65).Left = 0
  aChars(65).Width = CHAR_WIDE
  For x = 66 To 72
    aChars(x).Left = aChars(x - 1).Left + aChars(x - 1).Width
    aChars(x).Width = CHAR_WIDE
  Next x
  aChars(73).Left = aChars(72).Left + aChars(72).Width
  aChars(73).Width = CHAR_THIN
  aChars(74).Left = aChars(73).Left + aChars(73).Width
  aChars(74).Width = CHAR_WIDE
  aChars(75).Left = aChars(74).Left + aChars(74).Width
  aChars(75).Width = CHAR_THIN
  For x = 76 To 83
    aChars(x).Left = aChars(x - 1).Left + aChars(x - 1).Width
    aChars(x).Width = CHAR_WIDE
  Next x
  aChars(84).Left = aChars(83).Left + aChars(83).Width
  aChars(84).Width = CHAR_THIN
  aChars(85).Left = aChars(84).Left + aChars(84).Width
  aChars(85).Width = CHAR_WIDE
  aChars(86).Left = aChars(85).Left + aChars(85).Width
  aChars(86).Width = CHAR_THIN
  aChars(87).Left = aChars(86).Left + aChars(86).Width
  aChars(87).Width = 50
  For x = 88 To 90
    aChars(x).Left = aChars(x - 1).Left + aChars(x - 1).Width
    aChars(x).Width = CHAR_THIN
  Next x
End Sub

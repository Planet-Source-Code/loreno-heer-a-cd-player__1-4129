Attribute VB_Name = "Module2"
Private Declare Function PlaySound Lib "winmm.dll" Alias "PlaySoundA" (ByVal LpszName As String, ByVal hModule As Long, ByVal dwFlags As Long) As Long
Private Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLenght As Long, ByVal hwndCallback As Long) As Long

Private Const SND_SYNC As Long = &H0
Private Const SND_ASYNC As Long = &H1
Private Const SND_NODEFAULT As Long = &H2
Private Const SND_LOOP As Long = &H8
Private Const SND_FILENAME As Long = &H20000
Private Const SND_DEFAULTPATH As String = "C:\Windows\media\"

Dim privPath As String
Dim privNoDefault As Boolean
Dim privAnzTracks As Integer

Option Explicit
Property Get cdLength()
Dim lngResult As Long
Dim strBuffer As String

lngResult = mciSendString("open cdaudio", "", 0, 0)
strBuffer = String$(256, vbNullChar)
lngResult = mciSendString("status cdaudio number of tracks", strBuffer, Len(strBuffer), 0)
privAnzTracks = Val(Left$(strBuffer, InStr(strBuffer, vbNullChar) - 1))
lngResult = mciSendString("close cd audio", "", 0, 0)
cdLength = privAnzTracks
End Property
Public Sub playCdFull()
Dim lngResult As Long

lngResult = mciSendString("open cdaudio", "", 0, 0)
lngResult = mciSendString("play cdaudio", "", 0, 0)
If lngResult > 0 Then
ShapedForm.AXMarquee2.Text = "Keine Audio CD eingelegt"
lngResult = mciSendString("stop cdaudio", "", 0, 0)
lngResult = mciSendString("close cdaudio", "", 0, 0)
End If
End Sub
Public Sub doorOpen()
Dim lngResult As Long
lngResult = mciSendString("open cdaudio", "", 0, 0)
lngResult = mciSendString("set cdaudio door open", "", 0, 0)
lngResult = mciSendString("close cdaudio", "", 0, 0)
End Sub
Public Sub doorClose()
Dim lngResult As Long
lngResult = mciSendString("open cdaudio", "", 0, 0)
lngResult = mciSendString("set cdaudio door closed", "", 0, 0)
lngResult = mciSendString("close cdaudio", "", 0, 0)
End Sub
Public Sub playCdTrack(track As Integer)
Dim lngResult As Long
lngResult = mciSendString("open cdaudio", "", 0, 0)
lngResult = mciSendString("set cdaudio time format tmsf", "", 0, 0)
If track < Module2.cdLength Then
    lngResult = mciSendString("play cdaudio from " & track & " to " & track + 1, "", 0, 0)
Else
    lngResult = mciSendString("play cdaudio from " & track, "", 0, 0)
End If
lngResult = mciSendString("close cdaudio", "", 0, 0)
End Sub
Public Sub stopCd()
Dim lngResult As Long
lngResult = mciSendString("stop cdaudio", "", 0, 0)
lngResult = mciSendString("close cdaudio", "", 0, 0)
End Sub
Public Sub playCdShuffle()
Dim shuffle As Integer
Randomize Timer
shuffle = Int((Rnd * Module2.cdLength) + 1)
playCdTrack (shuffle)
End Sub
Public Sub mute()
Dim lngResult As Long
lngResult = mciSendString("set cdaudio audio all off", "", 0, 0)
End Sub
Public Sub DeMute()
Dim lngResult As Long
lngResult = mciSendString("set cdaudio audio all on", "", 0, 0)
End Sub

Attribute VB_Name = "Module1"
Option Explicit

Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public Declare Function playa Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long

Dim SND_FLAG As Long

Public Sub PlayWav(sFile As String)
    If Dir(sFile$) <> "" Then Call playa(sFile, SND_FLAG)
End Sub


Attribute VB_Name = "mdlSoundManager"
Option Explicit
Private Declare Function sndPlaySound Lib "winmm" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long

Public Enum SND_Settings
    SND_SYNC = &H0
    SND_ASYNC = &H1
    SND_NODEFAULT = &H2
    SND_MEMORY = &H4
    SND_LOOP = &H8
    SND_NOSTOP = &H10
    SW_SHOW = 5
End Enum

Sub PlayWave(strFilename As String, Optional Settings As SND_Settings = SND_ASYNC)
    Dim retval As Long
    retval = sndPlaySound(strFilename, Settings)
End Sub

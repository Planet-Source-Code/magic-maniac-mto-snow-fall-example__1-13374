Attribute VB_Name = "MSysSnd"
Option Explicit

Public Declare Function PlaySound Lib "winmm.dll" Alias "PlaySoundA" (ByVal lpszName As String, ByVal hModule As Long, ByVal dwFlags As Long) As Long
Public Declare Function StopSound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszNull As Long, ByVal uFlags As Long) As Long

Public Sub PlaySoundFile(ByVal FileName As String, Optional ByVal Wait As Boolean = False)
  If Wait Then
    Call PlaySound(FileName, 0&, &H20000)
  Else
    Call PlaySound(FileName, 0&, &H1 Or &H20000)
  End If
End Sub

Public Sub StopSoundFile()
  StopSound 0, 0
End Sub


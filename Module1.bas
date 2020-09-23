Attribute VB_Name = "Module1"
Type SFType
  X1 As Long
  Y1 As Long
  S1 As Long
  S2 As Long
  W1 As Long
  W2 As Long
  
  Ox1 As Long
  Oy1 As Long
End Type

Declare Function ShellExecute% Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd%, ByVal lpOperation$, ByVal lpFile$, ByVal lpParameters$, ByVal lpDirectory$, ByVal nShowCmd%)

Function Wait(Optional Seconds As Variant = 5)
  Dim TimerStarted As Long
  TimerStarted = Abs(86400 - Timer)
  Do While Abs(TimerStarted - ((86400 - Timer) + 1)) < Seconds
    DoEvents
  Loop
End Function

Function Url(sUrl As String) As Long
  Url = ShellExecute(0, "OPEN", sUrl$, vbNull, vbNull, vbNull)
End Function


VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MORTAL OBSESSiON - Wish you all a merry x-mas..."
   ClientHeight    =   3765
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5790
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3765
   ScaleWidth      =   5790
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   45
      Top             =   3825
   End
   Begin VB.PictureBox Pic1 
      ForeColor       =   &H000000FF&
      Height          =   3750
      Left            =   0
      ScaleHeight     =   246
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   381
      TabIndex        =   0
      Top             =   0
      Width           =   5775
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'# -------------------------------------
'# X-Mas Greet!
'#
'# Version 1.0
'# -------------------------------------
'# CODED BY:
'#
'# MAGiC MANiAC^mTo ( mto@kabelfoon.nl )
'#
'# MORTAL OBSESSiON
'# http://home.kabelfoon.nl/~mto
'# -------------------------------------
'# RELEASED 6-Dec-2000 ON:
'#
'# www.planet-source-code.com
'# -------------------------------------

Dim SFTotal As Long
Dim SF(500) As SFType
Dim lTmr As Long

Private Sub Form_Load()
  Dim lTmp1 As Long
  Dim lTmp2 As Long
  Dim lTmp3 As Long
  Dim bbol1 As Boolean
  
  Randomize
  
  SFTotal = 200 ' total snowflakes...
  
  If SFTotal > 500 Then
    SFTotal = 500
  End If
  For lTmp1 = 0 To SFTotal - 1
    SetSF lTmp1
  Next

  Load_Picture

  Me.Show
  
  Do
  For lTmp1 = 0 To SFTotal - 1
    
    If SF(lTmp1).W2 < SF(lTmp1).W1 Then
      SF(lTmp1).W2 = SF(lTmp1).W2 + 1
    Else
      If SF(lTmp1).S2 < SF(lTmp1).S1 Then
        SF(lTmp1).S2 = SF(lTmp1).S2 + 1
      Else
        SF(lTmp1).S2 = 0
        If SF(lTmp1).Y1 < Pic1.Height / Screen.TwipsPerPixelY - 5 Then
          
          If Pic1.Point(SF(lTmp1).X1, SF(lTmp1).Y1 + 1) <> RGB(0, 0, 0) Then
            
            lTmp2 = Rnd() * 1
          

            bbol1 = False
Again1:
            If lTmp2 = 0 Then
              If Pic1.Point(SF(lTmp1).X1 + 1, SF(lTmp1).Y1 + 1) = RGB(0, 0, 0) Then
                Pic1.PSet (SF(lTmp1).Ox1, SF(lTmp1).Oy1), RGB(0, 0, 0)
                SF(lTmp1).X1 = SF(lTmp1).X1 + 1
              Else
                lTmp2 = 1
              End If
            End If
            
            If lTmp2 = 1 Then
              If Pic1.Point(SF(lTmp1).X1 - 1, SF(lTmp1).Y1 + 1) = RGB(0, 0, 0) Then
                Pic1.PSet (SF(lTmp1).Ox1, SF(lTmp1).Oy1), RGB(0, 0, 0)
                SF(lTmp1).X1 = SF(lTmp1).X1 - 1
              Else
                If bbol1 = False Then
                  bbol1 = True
                  GoTo Again1
                End If
                lTmp2 = 2
              End If
            End If
          
              
            If lTmp2 = 2 Then
              SetSF lTmp1
              GoTo next1
            End If
          Else
            Pic1.PSet (SF(lTmp1).Ox1, SF(lTmp1).Oy1), RGB(0, 0, 0)
            SF(lTmp1).Y1 = SF(lTmp1).Y1 + 1
          End If
          
          SF(lTmp1).Ox1 = SF(lTmp1).X1
          SF(lTmp1).Oy1 = SF(lTmp1).Y1
          Pic1.PSet (SF(lTmp1).X1, SF(lTmp1).Y1), RGB(254, 254, 254)
        
          lTmp3 = Rnd() * 2 - 1
          If SF(lTmp1).X1 + lTmp3 < 1 Then
            lTmp3 = 1
          Else
            If SF(lTmp1).X1 + lTmp3 > Me.Pic1.Width - 1 Then
              lTmp3 = -1
            End If
          End If
          
          SF(lTmp1).X1 = SF(lTmp1).X1 + lTmp3
        
        Else
          SetSF lTmp1
        End If
      End If
    End If
next1:
    DoEvents
  Next
Loop


End Sub

Private Sub Form_Unload(Cancel As Integer)
  MsgBox "Vote This Code For X-Mas As A Present... ;)" & vbCrLf & vbCrLf & "www.planet-source-code.com"
  End
End Sub

Private Sub SetSF(SFNr)
  SF(SFNr).X1 = Rnd() * Pic1.Width / Screen.TwipsPerPixelX
  SF(SFNr).Y1 = 0
  SF(SFNr).S1 = Rnd() * 2
  SF(SFNr).S2 = 0
  SF(SFNr).W1 = Rnd() * Pic1.Height / Screen.TwipsPerPixelY
  SF(SFNr).W2 = 0
  SF(SFNr).Ox1 = -1
  SF(SFNr).Oy1 = -1
End Sub

Private Sub Pic1_Click()
  Url "http://home.kabelfoon.nl/~mto"
End Sub

Private Sub Timer1_Timer()
  lTmr = lTmr + 1
  If lTmr > 2000 Then
    Load_Picture
    lTmr = 0
  End If
End Sub

Sub Load_Picture()
  On Error Resume Next
  Me.Pic1.Picture = LoadPicture(App.Path + "\pic1.gif", , , 0, 0)
  If Err Then
    MsgBox "X-Mas Picture Was Not Found!"
    End
  End If
  On Error GoTo 0
  Pic1.Move 0, 0, Pic1.Width, Pic1.Height
  Me.Width = Pic1.Width
  Me.Height = Pic1.Height + 340
  Pic1.AutoSize = True
  Pic1.DrawWidth = 1
  Pic1.ForeColor = RGB(254, 254, 254)
  
  PlaySoundFile App.Path + "\" & LTrim((Int(Rnd() * 2)) + 1) & ".wav"
End Sub

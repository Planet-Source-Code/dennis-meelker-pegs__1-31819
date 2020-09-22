VERSION 5.00
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Pegs"
   ClientHeight    =   3450
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   3435
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "frmMain.frx":08CA
   ScaleHeight     =   230
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   229
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picTags 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   300
      Left            =   840
      Picture         =   "frmMain.frx":1194
      ScaleHeight     =   20
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   60
      TabIndex        =   0
      Top             =   3840
      Visible         =   0   'False
      Width           =   900
   End
   Begin VB.Menu mnuNew 
      Caption         =   "&New Game"
   End
   Begin VB.Menu mnuAbout 
      Caption         =   "&About"
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long

Private Const SND_ASYNC = &H1         '  play asynchronously
Private Const SND_NODEFAULT = &H2         '  silence not default, if sound not found

Private Const SRCAND = &H8800C6  ' (DWORD) dest = source AND dest
Private Const SRCCOPY = &HCC0020 ' (DWORD) dest = source
Private Const SRCPAINT = &HEE0086        ' (DWORD) dest = source OR dest

Dim iFrom As Integer
Dim bWinner As Boolean
Dim bBusy As Boolean

Private Hok(0 To 44) As HokType

Private Type HokType
    Top As Variant
    Left As Variant
    Tag As Variant
End Type

Private Sub Form_Load()
frmMain.Picture = LoadPicture(FixPath(App.Path) & "Wood.jpg")
StartNewGame
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Integer
If bWinner = True Then Exit Sub

i = GetTagNumber(X, Y)
If i > -1 Then
    iFrom = i
    sndPlaySound FixPath(App.Path) & "PickUp.wav", SND_ASYNC + SND_NODEFAULT
    frmMain.MousePointer = 99
End If
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Integer
If bWinner = True Then Exit Sub

'Y = Y + 13

frmMain.MousePointer = 0

i = GetTagNumber(X, Y)
If i = -1 Then Exit Sub
X = Hok(i).Left + 10
Y = Hok(i).Top + 10
If i > -1 Then
    If GetTagNumber(X + 40, Y) = iFrom Then
        If Hok(GetTagNumber(X + 20, Y)).Tag = 1 And Hok(i).Tag = 0 Then
            Hok(GetTagNumber(X + 20, Y)).Tag = 0
            Hok(i).Tag = 1
            Hok(iFrom).Tag = 0
            DrawTag iFrom, False
            DrawTag i, True
            DrawTag GetTagNumber(X + 20, Y), False
            sndPlaySound FixPath(App.Path) & "Zet.wav", SND_ASYNC + SND_NODEFAULT
        End If
    ElseIf GetTagNumber(X - 40, Y) = iFrom Then
        If Hok(GetTagNumber(X - 20, Y)).Tag = 1 And Hok(i).Tag = 0 Then
            Hok(GetTagNumber(X - 20, Y)).Tag = 0
            Hok(i).Tag = 1
            Hok(iFrom).Tag = 0
            DrawTag iFrom, False
            DrawTag i, True
            DrawTag GetTagNumber(X - 20, Y), False
            sndPlaySound FixPath(App.Path) & "Zet.wav", SND_ASYNC + SND_NODEFAULT
        End If
    ElseIf GetTagNumber(X, Y - 40) = iFrom Then
        If Hok(GetTagNumber(X, Y - 20)).Tag = 1 And Hok(i).Tag = 0 Then
            Hok(GetTagNumber(X, Y - 20)).Tag = 0
            Hok(i).Tag = 1
            Hok(iFrom).Tag = 0
            DrawTag iFrom, False
            DrawTag i, True
            DrawTag GetTagNumber(X, Y - 20), False
            sndPlaySound FixPath(App.Path) & "Zet.wav", SND_ASYNC + SND_NODEFAULT
        End If
    ElseIf GetTagNumber(X, Y + 40) = iFrom Then
        If Hok(GetTagNumber(X, Y + 20)).Tag = 1 And Hok(i).Tag = 0 Then
            Hok(GetTagNumber(X, Y + 20)).Tag = 0
            Hok(i).Tag = 1
            Hok(iFrom).Tag = 0
            DrawTag iFrom, False
            DrawTag i, True
            DrawTag GetTagNumber(X, Y + 20), False
            sndPlaySound FixPath(App.Path) & "Zet.wav", SND_ASYNC + SND_NODEFAULT
        End If
    End If
    
    CheckForWinner
    
    
    frmMain.Refresh
End If
End Sub

Public Function DrawTag(iIndex As Integer, bUsed As Boolean)
If bUsed = False Then
    BitBlt frmMain.hDC, Hok(iIndex).Left, Hok(iIndex).Top, 20, 20, picTags.hDC, 40, 0, SRCAND
    BitBlt frmMain.hDC, Hok(iIndex).Left, Hok(iIndex).Top, 20, 20, picTags.hDC, 20, 0, SRCPAINT
Else
    BitBlt frmMain.hDC, Hok(iIndex).Left, Hok(iIndex).Top, 20, 20, picTags.hDC, 40, 0, SRCAND
    BitBlt frmMain.hDC, Hok(iIndex).Left, Hok(iIndex).Top, 20, 20, picTags.hDC, 0, 0, SRCPAINT
End If
End Function

Public Function GetTagNumber(X As Variant, Y As Variant) As Integer
For i = 0 To 44
    If X > Hok(i).Left And X < Hok(i).Left + 20 And Y > Hok(i).Top And Y < Hok(i).Top + 20 Then
        GetTagNumber = i
        Exit Function
    End If
Next i
GetTagNumber = -1
End Function

Public Sub CheckForWinner()
Dim bCenterGood As Boolean

bCenterGood = False

For i = 0 To 44
    If i = 40 Then
        If Hok(40).Tag = 1 Then
            bCenterGood = True
        End If
    Else
        If Hok(i).Tag = 1 Then
            Exit Sub
        End If
    End If
Next i

If bCenterGood = True Then
    bWinner = True
    frmMain.MousePointer = 0
    MsgBox "You won!"
End If
End Sub

Public Sub InitHok()
Hok(0).Top = 8
Hok(0).Left = 80
Hok(0).Tag = 1
Hok(1).Top = 8
Hok(1).Left = 104
Hok(1).Tag = 1
Hok(2).Top = 8
Hok(2).Left = 128
Hok(2).Tag = 1
Hok(3).Top = 32
Hok(3).Left = 80
Hok(3).Tag = 1
Hok(4).Top = 32
Hok(4).Left = 104
Hok(4).Tag = 1
Hok(5).Top = 32
Hok(5).Left = 128
Hok(5).Tag = 1
Hok(6).Top = 56
Hok(6).Left = 80
Hok(6).Tag = 1
Hok(7).Top = 56
Hok(7).Left = 104
Hok(7).Tag = 1
Hok(8).Top = 56
Hok(8).Left = 128
Hok(8).Tag = 1
Hok(9).Top = 80
Hok(9).Left = 8
Hok(9).Tag = 1
Hok(10).Top = 80
Hok(10).Left = 32
Hok(10).Tag = 1
Hok(11).Top = 80
Hok(11).Left = 56
Hok(11).Tag = 1
Hok(12).Top = 104
Hok(12).Left = 8
Hok(12).Tag = 1
Hok(13).Top = 104
Hok(13).Left = 32
Hok(13).Tag = 1
Hok(14).Top = 104
Hok(14).Left = 56
Hok(14).Tag = 1
Hok(15).Top = 128
Hok(15).Left = 8
Hok(15).Tag = 1
Hok(16).Top = 128
Hok(16).Left = 32
Hok(16).Tag = 1
Hok(17).Top = 128
Hok(17).Left = 56
Hok(17).Tag = 1
Hok(18).Top = 80
Hok(18).Left = 152
Hok(18).Tag = 1
Hok(19).Top = 80
Hok(19).Left = 176
Hok(19).Tag = 1
Hok(20).Top = 80
Hok(20).Left = 200
Hok(20).Tag = 1
Hok(21).Top = 104
Hok(21).Left = 152
Hok(21).Tag = 1
Hok(22).Top = 104
Hok(22).Left = 176
Hok(22).Tag = 1
Hok(23).Top = 104
Hok(23).Left = 200
Hok(23).Tag = 1
Hok(24).Top = 128
Hok(24).Left = 152
Hok(24).Tag = 1
Hok(25).Top = 128
Hok(25).Left = 176
Hok(25).Tag = 1
Hok(26).Top = 128
Hok(26).Left = 200
Hok(26).Tag = 1
Hok(27).Top = 152
Hok(27).Left = 80
Hok(27).Tag = 1
Hok(28).Top = 152
Hok(28).Left = 104
Hok(28).Tag = 1
Hok(29).Top = 152
Hok(29).Left = 128
Hok(29).Tag = 1
Hok(30).Top = 176
Hok(30).Left = 80
Hok(30).Tag = 1
Hok(31).Top = 176
Hok(31).Left = 104
Hok(31).Tag = 1
Hok(32).Top = 176
Hok(32).Left = 128
Hok(32).Tag = 1
Hok(33).Top = 200
Hok(33).Left = 80
Hok(33).Tag = 1
Hok(34).Top = 200
Hok(34).Left = 104
Hok(34).Tag = 1
Hok(35).Top = 200
Hok(35).Left = 128
Hok(35).Tag = 1
Hok(36).Top = 80
Hok(36).Left = 80
Hok(36).Tag = 1
Hok(37).Top = 80
Hok(37).Left = 104
Hok(37).Tag = 1
Hok(38).Top = 80
Hok(38).Left = 128
Hok(38).Tag = 1
Hok(39).Top = 104
Hok(39).Left = 80
Hok(39).Tag = 1
Hok(40).Top = 104
Hok(40).Left = 104
Hok(40).Tag = 0
Hok(41).Top = 104
Hok(41).Left = 128
Hok(41).Tag = 1
Hok(42).Top = 128
Hok(42).Left = 80
Hok(42).Tag = 1
Hok(43).Top = 128
Hok(43).Left = 104
Hok(43).Tag = 1
Hok(44).Top = 128
Hok(44).Left = 128
Hok(44).Tag = 1
End Sub

Public Sub StartNewGame()
bWinner = False
bBusy = False
InitHok
For i = 0 To 44
    If i = 40 Then
        BitBlt frmMain.hDC, Hok(i).Left, Hok(i).Top, 20, 20, picTags.hDC, 40, 0, SRCAND
        BitBlt frmMain.hDC, Hok(i).Left, Hok(i).Top, 20, 20, picTags.hDC, 20, 0, SRCPAINT
        Hok(40).Tag = 0
    Else
        BitBlt frmMain.hDC, Hok(i).Left, Hok(i).Top, 20, 20, picTags.hDC, 40, 0, SRCAND
        BitBlt frmMain.hDC, Hok(i).Left, Hok(i).Top, 20, 20, picTags.hDC, 0, 0, SRCPAINT
        Hok(i).Tag = 1
    End If
Next i
frmMain.Refresh
End Sub

Private Sub mnuAbout_Click()
frmAbout.Show vbModal
End Sub

Private Sub mnuNew_Click()
If MsgBox("Are you sure you want to start a new game?", vbQuestion + vbYesNo) = vbYes Then
    StartNewGame
End If
End Sub

Public Function FixPath(sPath As Variant) As String
If Right(sPath, 1) = "\" Then
    FixPath = sPath
Else
    FixPath = sPath & "\"
End If
End Function

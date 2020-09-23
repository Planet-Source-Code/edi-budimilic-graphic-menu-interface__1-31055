VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "GS Home Games - build 021230"
   ClientHeight    =   7200
   ClientLeft      =   2865
   ClientTop       =   1785
   ClientWidth     =   9600
   ForeColor       =   &H00000000&
   Icon            =   "Form1.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form1.frx":08CA
   ScaleHeight     =   480
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   640
   Begin VB.PictureBox StatusBar 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   2760
      Picture         =   "Form1.frx":34A4
      ScaleHeight     =   465
      ScaleWidth      =   9600
      TabIndex        =   5
      Top             =   7320
      Visible         =   0   'False
      Width           =   9630
   End
   Begin VB.PictureBox Pozadina 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   3030
      Left            =   840
      ScaleHeight     =   3000
      ScaleWidth      =   3000
      TabIndex        =   4
      Top             =   7800
      Visible         =   0   'False
      Width           =   3030
   End
   Begin VB.PictureBox Gumb 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   120
      Picture         =   "Form1.frx":851D
      ScaleHeight     =   465
      ScaleWidth      =   2550
      TabIndex        =   3
      Top             =   7320
      Visible         =   0   'False
      Width           =   2580
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   8880
      Top             =   8880
   End
   Begin VB.PictureBox InfoBox 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   6030
      Left            =   120
      Picture         =   "Form1.frx":AAC1
      ScaleHeight     =   6000
      ScaleWidth      =   5250
      TabIndex        =   2
      Top             =   8760
      Visible         =   0   'False
      Width           =   5280
   End
   Begin VB.PictureBox MenuPic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   7200
      Left            =   0
      ScaleHeight     =   478
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   638
      TabIndex        =   1
      Top             =   0
      Width           =   9600
   End
   Begin VB.PictureBox Back1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   990
      Left            =   120
      Picture         =   "Form1.frx":25602
      ScaleHeight     =   960
      ScaleWidth      =   960
      TabIndex        =   0
      Top             =   7800
      Visible         =   0   'False
      Width           =   990
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim LUPx, LUPy, LupYz, rndA, rndB, FPS, KLIK, Sos As Integer
Dim GSirina(1 To 10) As Integer
Dim XX, YY  ' as mouse position

Private Sub Form_Activate()
Sos = 6 'Speed of Selection (button)
LupYz = 0
rndB = 0
1
DoEvents
MenuPic.Cls ' Clear screen

For LUPy = 0 To 10 'First loop starts - Draw a background in tiles
For LUPx = 0 To 10
Call BitBlt(MenuPic.hdc, LUPx * 64, LUPy * 64 - LupYz, 66, 66, Back1.hdc, 0, 0, SRCCOPY) '*128 is default size
Call BitBlt(MenuPic.hdc, LUPx * 64, LUPy * 64 - LupYz, 66, 66, Back1.hdc, 0, 0, SRCCOPY)
Next LUPx
Next LUPy 'End of loop loop-a

Call BitBlt(MenuPic.hdc, 0, 447, 640, 31, StatusBar.hdc, 0, 0, SRCAND) 'Position StatusBar
Call TextOut(MenuPic.hdc, 400, 450, Format(Date, "dd.mm.yyyy") & "  " & Format(Time, "HH:mm:ss"), 20) 'Date & clock in StatusBar

Call TextOut(MenuPic.hdc, 15, 450, "About / Credits", 15) 'Draw's text / credits about us
Credits
If XX > 15 And XX < 175 And YY > 445 And YY < 500 Then CreditsOn = 1 Else CreditsOn = 0


MenuPic.ForeColor = &HFFC0C0
MenuPic.FontSize = 32: Call TextOut(Form1.MenuPic.hdc, 160, 15, "GS Home Games", 13): MenuPic.FontSize = 18 'Text na vrhu
MenuPic.ForeColor = vbWhite

Call BitBlt(MenuPic.hdc, 50, 80, GSirina(1), 31, Gumb.hdc, 0, 0, SRCERASE) 'Position of button
Call TextOut(MenuPic.hdc, 70, 83, "Snakey", 6)
If XX > 50 And XX < 210 And YY > 80 + (39 * 0) And YY < 110 + (39 * 0) Then GSirina(1) = GSirina(1) + Sos Else GSirina(1) = GSirina(1) - Sos: If GSirina(1) < 120 Then GSirina(1) = 120 Else
If GSirina(1) > 172 Then GSirina(1) = 172: Snakey1
If KLIK = 1 Then If XX > 50 And XX < 210 And YY > 80 + (39 * 0) And YY < 110 + (39 * 0) Then Call Shell(App.Path & "\Snakey\Snakey.EXE", 1): TimesStarted(1) = TimesStarted(1) + 1: SaveFajl: End 'Pokrece igru , doda jedan TimesStarted ,Snima u fajl i zatvara ovaj program

Call BitBlt(MenuPic.hdc, 50, 119, GSirina(2), 31, Gumb.hdc, 0, 0, SRCERASE)  'Position of button
Call TextOut(MenuPic.hdc, 71, 122, "Yambo", 5)
If XX > 50 And XX < 210 And YY > 80 + (39 * 1) And YY < 110 + (39 * 1) Then GSirina(2) = GSirina(2) + Sos Else GSirina(2) = GSirina(2) - Sos: If GSirina(2) < 120 Then GSirina(2) = 120 Else
If GSirina(2) > 172 Then GSirina(2) = 172: Yambo
If KLIK = 1 Then If XX > 50 And XX < 210 And YY > 80 + (39 * 1) And YY < 110 + (39 * 1) Then Call Shell(App.Path & "\Yambo\Yambo.EXE", 1): TimesStarted(2) = TimesStarted(2) + 1: SaveFajl: End 'Pokrece igru , doda jedan TimesStarted ,Snima u fajl i zatvara ovaj program

Call BitBlt(MenuPic.hdc, 50, 158, GSirina(3), 31, Gumb.hdc, 0, 0, SRCERASE) 'Position of button
Call TextOut(MenuPic.hdc, 79, 161, "Tetris", 6)
If XX > 50 And XX < 210 And YY > 80 + (39 * 2) And YY < 110 + (39 * 2) Then GSirina(3) = GSirina(3) + Sos Else GSirina(3) = GSirina(3) - Sos: If GSirina(3) < 120 Then GSirina(3) = 120 Else
If GSirina(3) > 172 Then GSirina(3) = 172

Call BitBlt(MenuPic.hdc, 50, 197, GSirina(4), 31, Gumb.hdc, 0, 0, SRCERASE) 'Position of button
Call TextOut(MenuPic.hdc, 77, 200, "Kerpo", 5)
If XX > 50 And XX < 210 And YY > 80 + (39 * 3) And YY < 110 + (39 * 3) Then GSirina(4) = GSirina(4) + Sos Else GSirina(4) = GSirina(4) - Sos: If GSirina(4) < 120 Then GSirina(4) = 120 Else
If GSirina(4) > 172 Then GSirina(4) = 172

Call BitBlt(MenuPic.hdc, 50, 236, GSirina(5), 31, Gumb.hdc, 0, 0, SRCERASE)  'Position of button
Call TextOut(MenuPic.hdc, 57, 239, "Membrain", 8)
If XX > 50 And XX < 210 And YY > 80 + (39 * 4) And YY < 110 + (39 * 4) Then GSirina(5) = GSirina(5) + Sos Else GSirina(5) = GSirina(5) - Sos: If GSirina(5) < 120 Then GSirina(5) = 120 Else
If GSirina(5) > 172 Then GSirina(5) = 172

Call BitBlt(MenuPic.hdc, 50, 275, GSirina(6), 31, Gumb.hdc, 0, 0, SRCERASE) 'Position of button
Call TextOut(MenuPic.hdc, 74, 278, "Fillbox", 7)
If XX > 50 And XX < 210 And YY > 80 + (39 * 5) And YY < 110 + (39 * 5) Then GSirina(6) = GSirina(6) + Sos Else GSirina(6) = GSirina(6) - Sos: If GSirina(6) < 120 Then GSirina(6) = 120 Else
If GSirina(6) > 172 Then GSirina(6) = 172

Call BitBlt(MenuPic.hdc, 50, 314, GSirina(7), 31, Gumb.hdc, 0, 0, SRCERASE) 'Position of button
Call TextOut(MenuPic.hdc, 61, 317, "4 In Row", 8)
If XX > 50 And XX < 210 And YY > 80 + (39 * 6) And YY < 110 + (39 * 6) Then GSirina(7) = GSirina(7) + Sos Else GSirina(7) = GSirina(7) - Sos: If GSirina(7) < 120 Then GSirina(7) = 120 Else
If GSirina(7) > 172 Then GSirina(7) = 172: FourinRow

Call BitBlt(MenuPic.hdc, 50, 353, GSirina(8), 31, Gumb.hdc, 0, 0, SRCERASE)  'Position of button
Call TextOut(MenuPic.hdc, 54, 356, "Shipwreck", 9)
If XX > 50 And XX < 210 And YY > 80 + (39 * 7) And YY < 110 + (39 * 7) Then GSirina(8) = GSirina(8) + Sos Else GSirina(8) = GSirina(8) - Sos: If GSirina(8) < 120 Then GSirina(8) = 120 Else
If GSirina(8) > 172 Then GSirina(8) = 172


rndA = Rnd(1): rndA = rndA \ 1 'Random move of background
If rndA = 0 Then rndB = rndB - 1 Else rndB = rndB + 1
LupYz = LupYz - rndB
If LupYz < 0 Then LupYz = 128
If LupYz > 128 Then LupYz = 0
If rndB > 4 Then rndB = rndB - 1 'speed limit of random moving background
If rndB < -4 Then rndB = rndB + 1  'speed limit of random moving background

FPS = FPS + 1
GoTo 1
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then End ' Exit program
End Sub

Private Sub Form_Load()
Form1.Width = 9700
Form1.Height = 7600
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub MenuPic_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
KLIK = 1
End Sub

Private Sub MenuPic_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
XX = X
YY = Y
End Sub

Private Sub MenuPic_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
KLIK = 0
End Sub

Private Sub Timer1_Timer()
'Who gives a **** about GetTickCount
Form1.Caption = "FPS = " & FPS
FPS = 0
End Sub

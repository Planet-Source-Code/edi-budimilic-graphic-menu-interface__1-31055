Attribute VB_Name = "Module1"
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Const SRCCOPY = &HCC0020
Public Const SRCPAINT = &HEE0086
Public Const SRCAND = &H8800C6
Public Const SRCINVERT = &H660046
Public Const SRCERASE = &H440328
Public Const SND_SYNC = &H0
Public Const SND_ASYNC = &H1

Public Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Public TimesStarted(1 To 10), CreditsOn
Public InfoT, InfoH As Integer

Public Sub Credits()
'H = 368     T=80
If CreditsOn = 1 Then
InfoT = InfoT - 6
InfoH = InfoH + 6
If 448 + InfoT < 74 Then InfoT = InfoT + 6: InfoH = InfoH - 6
Call BitBlt(Form1.MenuPic.hdc, 220, 80, 350, 368 + InfoT, Form1.InfoBox.hdc, 0, 0, SRCERASE) 'Smedi
Call BitBlt(Form1.MenuPic.hdc, 220, 448 + InfoT, 350, -1 + InfoH, Form1.InfoBox.hdc, 0, 0, SRCAND) 'Plavi Info Box

If 448 + InfoT < 78 Then Krediti 'i pokrene kredite textove
Else
If 448 + InfoT < 448 Then InfoT = InfoT + 6: InfoH = InfoH - 6
Call BitBlt(Form1.MenuPic.hdc, 220, 80, 350, 368 + InfoT, Form1.InfoBox.hdc, 0, 0, SRCERASE) 'Smedi
Call BitBlt(Form1.MenuPic.hdc, 220, 448 + InfoT, 350, -1 + InfoH, Form1.InfoBox.hdc, 0, 0, SRCAND) 'Plavi Info Box
End If

End Sub
Public Sub LoadFajl()
Open App.Path & "\Counter.txt" For Input As #1
Input #1, TimesStarted(1)
Input #1, TimesStarted(2)
Input #1, TimesStarted(3)
Input #1, TimesStarted(4)
Input #1, TimesStarted(5)
Input #1, TimesStarted(6)
Input #1, TimesStarted(7)
Input #1, TimesStarted(8)
Input #1, TimesStarted(9)
Input #1, TimesStarted(10)
Close #1
End Sub
Public Sub SaveFajl()
Open App.Path & "\Counter.txt" For Output As #1
Print #1, TimesStarted(1)
Print #1, TimesStarted(2)
Print #1, TimesStarted(3)
Print #1, TimesStarted(4)
Print #1, TimesStarted(5)
Print #1, TimesStarted(6)
Print #1, TimesStarted(7)
Print #1, TimesStarted(8)
Print #1, TimesStarted(9)
Print #1, TimesStarted(10)
Close #1
End Sub
Public Sub Yambo()
On Error Resume Next
LoadFajl 'Ocitava brojac koliko se koj program puta pokrenuo

If Not Form1.Pozadina.Tag = "slika3" Then Form1.Pozadina.Picture = LoadPicture(App.Path & "\Yambo.jpg"): Form1.Pozadina.Tag = "slika3"
Form1.Pozadina.Height = 300 ' Samo zato sto je slika veca od drugih
Call BitBlt(Form1.MenuPic.hdc, 370, 149, 202, 300, Form1.Pozadina.hdc, 0, 0, SRCERASE) 'Pozadina/ScreenShot igrice
Form1.Pozadina.Height = 200 ' Samo zato sto je slika veca od drugih
Form1.MenuPic.FontSize = 18
Call TextOut(Form1.MenuPic.hdc, 225, 83, "Yambo", 5)
Form1.MenuPic.FontSize = 12
Call TextOut(Form1.MenuPic.hdc, 225, 118, "Minimum configuration : 133Mhz, 32Mb Ram", 40)
Call TextOut(Form1.MenuPic.hdc, 225, 148, "Multiplayer : None", 18)

Call TextOut(Form1.MenuPic.hdc, 225, 178, "Version : ", 10)
Call TextOut(Form1.MenuPic.hdc, 225, 198, "31.12.2001", 10)
Call TextOut(Form1.MenuPic.hdc, 225, 218, "1.02.03", 7)

Call TextOut(Form1.MenuPic.hdc, 225, 248, "Made by :", 9)
Call TextOut(Form1.MenuPic.hdc, 225, 268, "Edi Budimilic", 13)
Call TextOut(Form1.MenuPic.hdc, 225, 288, "Goran Duskic", 12)

Call TextOut(Form1.MenuPic.hdc, 225, 318, "No.of times started :", 21)
Call TextOut(Form1.MenuPic.hdc, 225, 338, TimesStarted(2), Len(TimesStarted(2)))

Form1.MenuPic.FontSize = 18

End Sub

Public Sub FourinRow()
On Error Resume Next
LoadFajl 'Ocitava brojac koliko se koj program puta pokrenuo

If Not Form1.Pozadina.Tag = "slika2" Then Form1.Pozadina.Picture = LoadPicture(App.Path & "\4inrow.jpg"): Form1.Pozadina.Tag = "slika2"
Call BitBlt(Form1.MenuPic.hdc, 370, 247, 202, 202, Form1.Pozadina.hdc, 0, 0, SRCERASE) 'Pozadina/ScreenShot igrice
Form1.MenuPic.FontSize = 18
Call TextOut(Form1.MenuPic.hdc, 225, 83, "4 In Row", 8)
Form1.MenuPic.FontSize = 12
Call TextOut(Form1.MenuPic.hdc, 225, 118, "Minimum configuration : 333Mhz, 32Mb Ram", 40)
Call TextOut(Form1.MenuPic.hdc, 225, 148, "Multiplayer : 2 players on one computer", 39)

Call TextOut(Form1.MenuPic.hdc, 225, 178, "Version : ", 10)
Call TextOut(Form1.MenuPic.hdc, 225, 198, "28.12.2001", 10)
Call TextOut(Form1.MenuPic.hdc, 225, 218, "0.03.00b", 8)

Call TextOut(Form1.MenuPic.hdc, 225, 248, "Made by :", 9)
Call TextOut(Form1.MenuPic.hdc, 225, 268, "Edi Budimilic", 13)
Call TextOut(Form1.MenuPic.hdc, 225, 288, "Goran Duskic", 12)

Call TextOut(Form1.MenuPic.hdc, 225, 318, "No.of times started :", 21)
Call TextOut(Form1.MenuPic.hdc, 225, 338, TimesStarted(7), Len(TimesStarted(7)))

Form1.MenuPic.FontSize = 18
End Sub

Public Sub Snakey1()
On Error Resume Next
LoadFajl 'Ocitava brojac koliko se koj program puta pokrenuo

If Not Form1.Pozadina.Tag = "slika1" Then Form1.Pozadina.Picture = LoadPicture(App.Path & "\Snakey 1.jpg"): Form1.Pozadina.Tag = "slika1"
Call BitBlt(Form1.MenuPic.hdc, 370, 247, 202, 202, Form1.Pozadina.hdc, 0, 0, SRCERASE) 'Pozadina/ScreenShot igrice
Form1.MenuPic.FontSize = 18
Call TextOut(Form1.MenuPic.hdc, 225, 83, "Snakey 1", 8)
Form1.MenuPic.FontSize = 12
Call TextOut(Form1.MenuPic.hdc, 225, 118, "Minimum configuration : 166Mhz, 32Mb Ram", 40)
Call TextOut(Form1.MenuPic.hdc, 225, 148, "Multiplayer : None", 18)

Call TextOut(Form1.MenuPic.hdc, 225, 178, "Version : ", 10)
Call TextOut(Form1.MenuPic.hdc, 225, 198, "31.08.2001", 10)
Call TextOut(Form1.MenuPic.hdc, 225, 218, "1.32.34", 7)

Call TextOut(Form1.MenuPic.hdc, 225, 248, "Made by :", 9)
Call TextOut(Form1.MenuPic.hdc, 225, 268, "Edi Budimilic", 13)
Call TextOut(Form1.MenuPic.hdc, 225, 288, "Goran Duskic", 12)

Call TextOut(Form1.MenuPic.hdc, 225, 318, "No.of times started :", 21)
Call TextOut(Form1.MenuPic.hdc, 225, 338, TimesStarted(1), Len(TimesStarted(1)))

Form1.MenuPic.FontSize = 18
End Sub
Public Sub Krediti()
On Error Resume Next
LoadFajl 'Ocitava brojac koliko se koj program puta pokrenuo

''If Not Form1.Pozadina.Tag = "slika1" Then Form1.Pozadina.Picture = LoadPicture(App.Path & "\Snakey 1.jpg"): Form1.Pozadina.Tag = "slika1"
''Call BitBlt(Form1.MenuPic.hdc, 370, 247, 202, 202, Form1.Pozadina.hdc, 0, 0, SRCERASE) 'Pozadina/ScreenShot igrice
Form1.MenuPic.FontSize = 18: Form1.MenuPic.ForeColor = &HFFC0C0
Call TextOut(Form1.MenuPic.hdc, 225, 83, "Generation Stars Team", 21)
Form1.MenuPic.FontSize = 12: Form1.MenuPic.ForeColor = vbWhite

Form1.MenuPic.FontBold = True: Call TextOut(Form1.MenuPic.hdc, 225, 120, "Edi Budimilic", 13)
Form1.MenuPic.FontBold = False: Call TextOut(Form1.MenuPic.hdc, 225, 140, "grommm@yahoo.com", 16)
Call TextOut(Form1.MenuPic.hdc, 225, 160, "http:\\bigbear.cjb.net", 22)

Form1.MenuPic.FontBold = True: Call TextOut(Form1.MenuPic.hdc, 225, 200, "Goran Duskic", 12)
Form1.MenuPic.FontBold = False: Call TextOut(Form1.MenuPic.hdc, 225, 220, "firedule@yahoo.com", 18)
Call TextOut(Form1.MenuPic.hdc, 225, 240, "http:\\firedule.cjb.net", 23)

Form1.MenuPic.FontBold = True: Call TextOut(Form1.MenuPic.hdc, 225, 280, "Dario Bognolo", 13)
Form1.MenuPic.FontBold = False: Call TextOut(Form1.MenuPic.hdc, 225, 300, "darhr@yahoo.com", 15)

Form1.MenuPic.FontBold = True: Call TextOut(Form1.MenuPic.hdc, 225, 340, "Dragan Maurovic", 15)
Form1.MenuPic.FontBold = False: Call TextOut(Form1.MenuPic.hdc, 225, 360, "biohazardnod@yahoo.com", 22)

Form1.MenuPic.ForeColor = &HFFC0C0: Form1.MenuPic.FontBold = True: Call TextOut(Form1.MenuPic.hdc, 225, 410, "http:\\generationstars.cjb.net", 30)
Form1.MenuPic.FontBold = False: Form1.MenuPic.ForeColor = vbWhite

Form1.MenuPic.FontSize = 18
End Sub

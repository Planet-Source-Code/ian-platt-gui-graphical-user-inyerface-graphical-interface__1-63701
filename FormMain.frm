VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form FormMain 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   12570
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   19095
   Icon            =   "FormMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   12570
   ScaleWidth      =   19095
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture23 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1260
      Left            =   11040
      Picture         =   "FormMain.frx":08CA
      ScaleHeight     =   1260
      ScaleWidth      =   1215
      TabIndex        =   29
      ToolTipText     =   "Play Video"
      Top             =   2520
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox Picture22 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1260
      Left            =   11040
      Picture         =   "FormMain.frx":591E
      ScaleHeight     =   1260
      ScaleWidth      =   1215
      TabIndex        =   28
      ToolTipText     =   "Play Video"
      Top             =   960
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox Picture19 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6375
      Left            =   18120
      Picture         =   "FormMain.frx":A972
      ScaleHeight     =   6375
      ScaleWidth      =   6345
      TabIndex        =   26
      ToolTipText     =   "Play Video"
      Top             =   120
      Visible         =   0   'False
      Width           =   6345
   End
   Begin VB.PictureBox Picture20 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6375
      Left            =   18120
      Picture         =   "FormMain.frx":E8EB
      ScaleHeight     =   6375
      ScaleWidth      =   6345
      TabIndex        =   25
      ToolTipText     =   "Play Video"
      Top             =   6720
      Visible         =   0   'False
      Width           =   6345
   End
   Begin VB.PictureBox Picture18 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6375
      Left            =   9480
      Picture         =   "FormMain.frx":12991
      ScaleHeight     =   6375
      ScaleWidth      =   6345
      TabIndex        =   24
      ToolTipText     =   "Play Video"
      Top             =   6600
      Visible         =   0   'False
      Width           =   6345
   End
   Begin VB.PictureBox Picture17 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1065
      Left            =   12240
      Picture         =   "FormMain.frx":167D0
      ScaleHeight     =   1065
      ScaleWidth      =   1800
      TabIndex        =   23
      ToolTipText     =   "Play Video"
      Top             =   8040
      Visible         =   0   'False
      Width           =   1800
   End
   Begin VB.PictureBox Picture16 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1080
      Left            =   10680
      Picture         =   "FormMain.frx":1CBEC
      ScaleHeight     =   1080
      ScaleWidth      =   1185
      TabIndex        =   22
      ToolTipText     =   "Play Video"
      Top             =   8040
      Visible         =   0   'False
      Width           =   1185
   End
   Begin VB.PictureBox Picture15 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1200
      Left            =   16200
      Picture         =   "FormMain.frx":20FB0
      ScaleHeight     =   1200
      ScaleWidth      =   1320
      TabIndex        =   21
      ToolTipText     =   "Open Graphics Program"
      Top             =   120
      Visible         =   0   'False
      Width           =   1320
   End
   Begin VB.PictureBox Picture14 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1080
      Left            =   16200
      Picture         =   "FormMain.frx":26274
      ScaleHeight     =   1080
      ScaleWidth      =   1125
      TabIndex        =   20
      ToolTipText     =   "Play Media Clips"
      Top             =   1560
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.PictureBox Picture13 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1065
      Left            =   16200
      Picture         =   "FormMain.frx":2A2D8
      ScaleHeight     =   1065
      ScaleWidth      =   1275
      TabIndex        =   19
      ToolTipText     =   "Browse Your Favourite Documents"
      Top             =   2760
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.PictureBox Picture10 
      AutoSize        =   -1  'True
      Height          =   1125
      Left            =   16200
      Picture         =   "FormMain.frx":2EA1C
      ScaleHeight     =   1065
      ScaleWidth      =   1800
      TabIndex        =   16
      Top             =   3960
      Width           =   1860
   End
   Begin VB.PictureBox Picture9 
      AutoSize        =   -1  'True
      Height          =   1140
      Left            =   16200
      Picture         =   "FormMain.frx":34E38
      ScaleHeight     =   1080
      ScaleWidth      =   1185
      TabIndex        =   15
      Top             =   5160
      Width           =   1245
   End
   Begin VB.PictureBox Picture8 
      AutoSize        =   -1  'True
      Height          =   1260
      Left            =   16200
      Picture         =   "FormMain.frx":391FC
      ScaleHeight     =   1200
      ScaleWidth      =   1320
      TabIndex        =   14
      Top             =   6360
      Width           =   1380
   End
   Begin VB.PictureBox Picture7 
      AutoSize        =   -1  'True
      Height          =   1140
      Left            =   16200
      Picture         =   "FormMain.frx":3E4C0
      ScaleHeight     =   1080
      ScaleWidth      =   1125
      TabIndex        =   13
      Top             =   7680
      Width           =   1185
   End
   Begin VB.PictureBox Picture6 
      AutoSize        =   -1  'True
      Height          =   1125
      Left            =   16200
      Picture         =   "FormMain.frx":42524
      ScaleHeight     =   1065
      ScaleWidth      =   1275
      TabIndex        =   12
      Top             =   8880
      Width           =   1335
   End
   Begin VB.Timer Timer6 
      Interval        =   100
      Left            =   120
      Top             =   6480
   End
   Begin VB.Timer Timer5 
      Interval        =   7000
      Left            =   120
      Top             =   6960
   End
   Begin VB.Timer Timer4 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   120
      Top             =   7440
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   120
      Top             =   7920
   End
   Begin VB.Timer Timer2 
      Interval        =   4500
      Left            =   120
      Top             =   8400
   End
   Begin VB.Timer Timer1 
      Interval        =   150
      Left            =   120
      Top             =   8880
   End
   Begin VB.PictureBox PicmainSkin 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6375
      Left            =   0
      Picture         =   "FormMain.frx":46C68
      ScaleHeight     =   6375
      ScaleWidth      =   6345
      TabIndex        =   0
      Top             =   0
      Width           =   6345
      Begin VB.PictureBox Picture21 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1260
         Left            =   4370
         Picture         =   "FormMain.frx":4AB2F
         ScaleHeight     =   1260
         ScaleWidth      =   1215
         TabIndex        =   27
         ToolTipText     =   "Play Video"
         Top             =   3840
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.PictureBox Picture5 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1065
         Left            =   4475
         Picture         =   "FormMain.frx":4FB83
         ScaleHeight     =   1065
         ScaleWidth      =   1800
         TabIndex        =   11
         ToolTipText     =   "Play Video"
         Top             =   2400
         Visible         =   0   'False
         Width           =   1800
      End
      Begin VB.PictureBox Picture4 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1080
         Left            =   2640
         Picture         =   "FormMain.frx":55F9F
         ScaleHeight     =   1080
         ScaleWidth      =   1185
         TabIndex        =   10
         ToolTipText     =   "Play Video"
         Top             =   4940
         Visible         =   0   'False
         Width           =   1185
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1080
         Left            =   460
         Picture         =   "FormMain.frx":5A363
         ScaleHeight     =   1080
         ScaleWidth      =   1125
         TabIndex        =   9
         ToolTipText     =   "Play Media Clips"
         Top             =   1435
         Visible         =   0   'False
         Width           =   1125
      End
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1200
         Left            =   500
         Picture         =   "FormMain.frx":5E3C7
         ScaleHeight     =   1200
         ScaleWidth      =   1320
         TabIndex        =   8
         ToolTipText     =   "Open Graphics Program"
         Top             =   3775
         Visible         =   0   'False
         Width           =   1320
      End
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   255
         Left            =   2040
         TabIndex        =   6
         Top             =   4800
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   0
         Scrolling       =   1
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1065
         Left            =   2525
         Picture         =   "FormMain.frx":6368B
         ScaleHeight     =   1065
         ScaleWidth      =   1275
         TabIndex        =   5
         ToolTipText     =   "Browse Your Favourite Documents"
         Top             =   300
         Visible         =   0   'False
         Width           =   1275
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "LOADING"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   2800
         TabIndex        =   7
         Top             =   5160
         Width           =   975
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00EDAC8B&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Impact"
            Size            =   27.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   2070
         TabIndex        =   1
         Top             =   1440
         Width           =   2235
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackColor       =   &H00EDAC8B&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Impact"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1755
         Left            =   2160
         TabIndex        =   4
         Top             =   2880
         Width           =   2175
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00EDAC8B&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Impact"
            Size            =   27.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   6735
         TabIndex        =   3
         Top             =   2880
         Width           =   135
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00EDAC8B&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Impact"
            Size            =   72
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1755
         Left            =   1080
         TabIndex        =   2
         Top             =   2400
         Width           =   255
      End
   End
   Begin VB.PictureBox Picture11 
      AutoSize        =   -1  'True
      Height          =   6435
      Left            =   0
      Picture         =   "FormMain.frx":67DCF
      ScaleHeight     =   6375
      ScaleWidth      =   9360
      TabIndex        =   17
      Top             =   6480
      Width           =   9420
   End
   Begin VB.PictureBox Picture12 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   6405
      Left            =   0
      Picture         =   "FormMain.frx":6C2D4
      ScaleHeight     =   6375
      ScaleWidth      =   9360
      TabIndex        =   18
      Top             =   0
      Width           =   9390
   End
End
Attribute VB_Name = "FormMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim framecount As Integer
Dim FRAMECOUNT2 As Integer
Dim TOPEND As Integer
Dim LEFTEND As Integer



Dim XX As Integer
'====================================================================
' Set up Screen Shape based on PicmainSkin
'====================================================================



Private Sub Form_Load()
    Dim WindowRegion As Long
    PicmainSkin.ScaleMode = vbPixels
    PicmainSkin.AutoRedraw = True
    PicmainSkin.AutoSize = True
    PicmainSkin.BorderStyle = vbBSNone
    Me.BorderStyle = vbBSNone
    'Set PicmainSkin.Picture = LoadPicture(App.Path & "\1.gif")
    Me.Width = PicmainSkin.Width
    Me.Height = PicmainSkin.Height
    WindowRegion = MakeRegion(PicmainSkin)
    SetWindowRgn Me.hWnd, WindowRegion, True
End Sub






'====================================================================
' Allow movement of form with mouse
'====================================================================

Private Sub PicmainSkin_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
      ReleaseCapture
      SendMessage Me.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End Sub

'====================================================================
' Allow Closure of form with Double Click
'====================================================================
Private Sub PicmainSkin_DblClick()
End
End Sub

Private Sub Picture1_Click()
On Error GoTo messageword
Shell "C:\Program Files\Microsoft Office\OFFICE11\WINWORD.EXE"
Exit Sub
messageword:
Label4.Caption = "Microsoft Word (2003) Not Installed"
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Set Picture1.Picture = Picture6.Picture 'LoadPicture(App.Path & "\prog1-2.bmp")
Me.Label4.Caption = "Word Processing"
End Sub

Private Sub Picture1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Set Picture1.Picture = Picture13.Picture ' LoadPicture(App.Path & "\prog1-1.bmp")
Me.Label4.Caption = ""
End Sub

Private Sub Picture2_Click()
On Error GoTo errormessage1

Shell "C:\Program Files\Windows Media Player\wmplayer.exe"
Exit Sub
errormessage1:
Label4.Caption = "Windows Media Player Not Installed"
End Sub

Private Sub Picture21_Click()
On Error GoTo errormessage21

Shell "C:\Program Files\Internet Explorer\iexplore.exe"
Exit Sub
errormessage21:
Label4.Caption = "Internet Explorer Not installed"
End Sub



Private Sub Picture2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Set Picture2.Picture = Picture7.Picture 'LoadPicture(App.Path & "\prog2-2.bmp")
Me.Label4.Caption = "Play Media"
End Sub

Private Sub Picture2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Set Picture2.Picture = Picture14.Picture  'LoadPicture(App.Path & "\prog2-1.bmp")
Me.Label4.Caption = ""
End Sub

Private Sub Picture3_Click()
On Error GoTo errormessage2
Shell "C:\Program Files\Adobe\Photoshop 7.0\Photoshop.exe"
Exit Sub
errormessage2:
Label4.Caption = "Adobe Photoshop 7 Not installed"

End Sub

Private Sub Picture3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Set Picture3.Picture = Picture8.Picture  ' LoadPicture(App.Path & "\prog3-2.bmp")
Me.Label4.Caption = "Graphics"
End Sub
Private Sub Picture3_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Set Picture3.Picture = Picture15.Picture ' LoadPicture(App.Path & "\prog3-1.bmp")
Me.Label4.Caption = ""
End Sub

Private Sub Picture4_Click()
On Error GoTo message3
Shell "C:\Program Files\Real\RealOne Player\realplay.exe"
Exit Sub
message3:
Label4.Caption = "Real Player One Not Installed"
End Sub

Private Sub Picture4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Set Picture4.Picture = Picture9.Picture 'LoadPicture(App.Path & "\prog4-2.bmp")
Me.Label4.Caption = "Play Video"
End Sub

Private Sub Picture4_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Set Picture4.Picture = Picture16.Picture 'LoadPicture(App.Path & "\prog4-1.bmp")
Me.Label4.Caption = ""
End Sub

Private Sub Picture5_Click()
End
End Sub

Private Sub Picture5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Set Picture5.Picture = Picture10.Picture 'LoadPicture(App.Path & "\exit2.bmp")
Me.Label4.Caption = "Exit"
End Sub

Private Sub Picture5_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Set Picture5.Picture = Picture17.Picture ' LoadPicture(App.Path & "\exit1.bmp")
Me.Label4.Caption = ""
End Sub

Private Sub Picture21_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Set Picture21.Picture = Picture23.Picture 'LoadPicture(App.Path & "\exit2.bmp")
Me.Label4.Caption = "Internet"
End Sub

Private Sub Picture21_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Set Picture21.Picture = Picture22.Picture ' LoadPicture(App.Path & "\exit1.bmp")
Me.Label4.Caption = ""
End Sub


'====================================================================
' Rotatae through pics from 1 to 11 and then back again
'====================================================================
Private Sub Timer1_Timer()
framecount = framecount + 1


If framecount = 1 Then Me.Label1.Caption = "G"
If framecount = 2 Then Me.Label1.Caption = "Gr"
If framecount = 3 Then Me.Label1.Caption = "Gra"
If framecount = 4 Then Me.Label1.Caption = "Grap"
If framecount = 5 Then Me.Label1.Caption = "Graph"
If framecount = 6 Then Me.Label1.Caption = "Graphi"
If framecount = 7 Then Me.Label1.Caption = "Graphic"
If framecount = 8 Then Me.Label1.Caption = "Graphica"
If framecount = 9 Then Me.Label1.Caption = "Graphical"

If framecount = 10 Then Me.Label2.Caption = "U"
If framecount = 11 Then Me.Label2.Caption = "US"
If framecount = 12 Then Me.Label2.Caption = "USE"
If framecount = 13 Then Me.Label2.Caption = "USER"
    
If framecount = 20 Then Set PicmainSkin.Picture = Picture18.Picture 'LoadPicture(App.Path & "\2.gif")
If framecount = 22 Then Set PicmainSkin.Picture = Picture19.Picture 'LoadPicture(App.Path & "\3.gif")
If framecount = 24 Then Set PicmainSkin.Picture = Picture20.Picture ' LoadPicture(App.Path & "\4.gif"): Timer1.Enabled = False


End Sub

Private Sub Timer2_Timer()
Timer2.Enabled = False
Timer3.Enabled = True
End Sub

Private Sub Timer3_Timer()

XX = XX + 1

Me.Left = Me.Left - 20

If XX = 100 Then Timer3.Enabled = False: Call panel1

End Sub

Private Sub panel1()
    Dim WindowRegion As Long
    PicmainSkin.ScaleMode = vbPixels
    PicmainSkin.AutoRedraw = True
    PicmainSkin.AutoSize = True
    PicmainSkin.BorderStyle = vbBSNone
    Me.BorderStyle = vbBSNone
    'Picture12.Top = 0: Picture12.Left = 0
    Set PicmainSkin.Picture = Picture11.Picture 'LoadPicture(App.Path & "\panel2.gif")
    Me.Width = PicmainSkin.Width
    Me.Height = PicmainSkin.Height
    WindowRegion = MakeRegion(PicmainSkin)
    SetWindowRgn Me.hWnd, WindowRegion, True
    Timer4.Enabled = True
End Sub



Private Sub Timer4_Timer()
FRAMECOUNT2 = FRAMECOUNT2 + 1


If FRAMECOUNT2 = 1 Then Me.Label3.Caption = "I"
If FRAMECOUNT2 = 2 Then Me.Label3.Caption = "IN"
If FRAMECOUNT2 = 3 Then Me.Label3.Caption = "INY"
If FRAMECOUNT2 = 4 Then Me.Label3.Caption = "INYE"
If FRAMECOUNT2 = 5 Then Me.Label3.Caption = "INYER"
If FRAMECOUNT2 = 6 Then Me.Label3.Caption = "INYERF"
If FRAMECOUNT2 = 7 Then Me.Label3.Caption = "INYERFA"
If FRAMECOUNT2 = 8 Then Me.Label3.Caption = "INYERFAC"
If FRAMECOUNT2 = 9 Then Me.Label3.Caption = "INYERFACE": Timer4.Enabled = False:

End Sub

Private Sub Timer5_Timer()


Call screen2: Timer5.Enabled = False



End Sub

Private Sub screen2()

Me.Label1.Caption = ""
Me.Label2.Caption = ""
Me.Label3.Caption = ""

Form1.Top = 0
Form1.Left = 0
Form1.Show

TOPEND = FormMain.Top + 2800
LEFTEND = FormMain.Left + 5400

For i = 1 To 100
Form1.Top = Form1.Top + TOPEND / 100
Form1.Left = Form1.Left + LEFTEND / 100
Next i
Form1.Hide
'Picture12.Top = 0: Picture12.Left = 0
    Set PicmainSkin.Picture = Picture12.Picture 'LoadPicture(App.Path & "\panel3.gif")

Picture1.Visible = True
Picture2.Visible = True
Picture3.Visible = True
Picture4.Visible = True
Picture5.Visible = True
Picture21.Visible = True

End Sub


Private Sub Timer6_Timer()
Me.ProgressBar1.Value = Me.ProgressBar1.Value + 2
Me.Label5.Caption = "Loading"

If Me.ProgressBar1.Value = 100 Then Timer6.Enabled = False: ProgressBar1.Visible = False: Me.Label5.Visible = False


End Sub

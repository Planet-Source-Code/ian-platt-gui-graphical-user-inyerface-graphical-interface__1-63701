VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picmainskin 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   900
      Left            =   0
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   900
      ScaleWidth      =   3015
      TabIndex        =   0
      Top             =   0
      Width           =   3015
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Dim WindowRegion As Long
    PicmainSkin.ScaleMode = vbPixels
    PicmainSkin.AutoRedraw = True
    PicmainSkin.AutoSize = True
    PicmainSkin.BorderStyle = vbBSNone
    Me.BorderStyle = vbBSNone
    'Set PicmainSkin.Picture = LoadPicture(App.Path & "\guilogo1.gif")
    Me.Width = PicmainSkin.Width
    Me.Height = PicmainSkin.Height
    WindowRegion = MakeRegion(PicmainSkin)
    SetWindowRgn Me.hWnd, WindowRegion, True
End Sub


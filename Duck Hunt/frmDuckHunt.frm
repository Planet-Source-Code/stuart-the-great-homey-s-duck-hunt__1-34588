VERSION 5.00
Begin VB.Form frmDuckHunt 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Duck Hunt-Stuart the Great"
   ClientHeight    =   6390
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6390
   ScaleWidth      =   7680
   StartUpPosition =   3  'Windows Default
   Begin VB.OptionButton optShotGun 
      Caption         =   "Shot Gun"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   240
      TabIndex        =   2
      Top             =   6000
      Value           =   -1  'True
      Width           =   1215
   End
   Begin VB.Frame fraGuns 
      Caption         =   "Guns"
      Height          =   615
      Left            =   0
      TabIndex        =   1
      Top             =   5760
      Width           =   4695
   End
   Begin VB.Image Image5 
      Height          =   1260
      Left            =   3480
      Picture         =   "frmDuckHunt.frx":0000
      Top             =   4920
      Visible         =   0   'False
      Width           =   1725
   End
   Begin VB.Image Image4 
      Height          =   1320
      Left            =   3480
      Picture         =   "frmDuckHunt.frx":0783
      Top             =   4800
      Visible         =   0   'False
      Width           =   1155
   End
   Begin VB.Image Image3 
      Height          =   1260
      Left            =   2640
      Picture         =   "frmDuckHunt.frx":0EAD
      Top             =   4920
      Visible         =   0   'False
      Width           =   1725
   End
   Begin VB.Image Image2 
      Height          =   1320
      Left            =   3120
      Picture         =   "frmDuckHunt.frx":1731
      Top             =   4800
      Visible         =   0   'False
      Width           =   1155
   End
   Begin VB.Image Image1 
      Height          =   1350
      Left            =   3360
      Picture         =   "frmDuckHunt.frx":1FBB
      Top             =   4800
      Visible         =   0   'False
      Width           =   930
   End
   Begin VB.Image imgDogJump 
      Height          =   1545
      Left            =   2280
      MousePointer    =   2  'Cross
      Picture         =   "frmDuckHunt.frx":2674
      Top             =   2640
      Visible         =   0   'False
      Width           =   1530
   End
   Begin VB.Image imgDogSniff 
      Height          =   1485
      Left            =   1320
      MousePointer    =   2  'Cross
      Picture         =   "frmDuckHunt.frx":310C
      Top             =   4080
      Width           =   1680
   End
   Begin VB.Image imgPlay 
      Height          =   2640
      Left            =   2640
      MousePointer    =   2  'Cross
      Picture         =   "frmDuckHunt.frx":417F
      Top             =   840
      Width           =   2445
   End
   Begin VB.Image imgHit 
      Height          =   1005
      Left            =   3240
      MousePointer    =   2  'Cross
      Picture         =   "frmDuckHunt.frx":4BA4
      Top             =   1800
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Image imgHaha 
      Height          =   1905
      Left            =   2520
      MousePointer    =   2  'Cross
      Picture         =   "frmDuckHunt.frx":5321
      Top             =   1320
      Visible         =   0   'False
      Width           =   2955
   End
   Begin VB.Label lblScore 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000B&
      Caption         =   "Score: 0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   4800
      TabIndex        =   0
      Top             =   5880
      Width           =   2895
   End
   Begin VB.Image imgDuckDie 
      Height          =   975
      Left            =   6600
      MousePointer    =   2  'Cross
      Picture         =   "frmDuckHunt.frx":6740
      Top             =   2520
      Visible         =   0   'False
      Width           =   930
   End
   Begin VB.Image imgDuckFly 
      Height          =   735
      Left            =   6720
      MousePointer    =   2  'Cross
      Picture         =   "frmDuckHunt.frx":716B
      Top             =   3480
      Visible         =   0   'False
      Width           =   900
   End
   Begin VB.Image imgDogDie 
      Height          =   1725
      Left            =   3120
      MousePointer    =   2  'Cross
      Picture         =   "frmDuckHunt.frx":781C
      Top             =   3240
      Visible         =   0   'False
      Width           =   1650
   End
   Begin VB.Image imgDogLaugh 
      Height          =   1590
      Left            =   3360
      MousePointer    =   2  'Cross
      Picture         =   "frmDuckHunt.frx":88C5
      Top             =   3360
      Visible         =   0   'False
      Width           =   1080
   End
   Begin VB.Image imgBack 
      Height          =   6540
      Left            =   0
      MousePointer    =   2  'Cross
      Picture         =   "frmDuckHunt.frx":951D
      Top             =   0
      Width           =   7680
   End
End
Attribute VB_Name = "frmDuckHunt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Dim intScore As Integer
    Dim intRandomLeft As Integer
    Dim intRandomTop As Integer
    Option Explicit
Private Sub Form_Load()
    intScore = 0
    imgBack.Height = 5745
End Sub



Private Sub imgBack_Click()
    If imgDogLaugh.Visible = True Then
        intRandomTop = Rnd * 3400
        intRandomLeft = Rnd * 6700
    
        imgDuckFly.Top = intRandomTop
        imgDuckFly.Left = intRandomLeft
    
        imgDuckFly.Visible = True
        imgDuckDie.Visible = False
        imgHaha.Visible = False
        imgDogLaugh.Visible = False
    ElseIf imgDogDie.Visible = True Then
        intRandomTop = Rnd * 3400
        intRandomLeft = Rnd * 6700
    
        imgDuckFly.Top = intRandomTop
        imgDuckFly.Left = intRandomLeft
    
        imgDuckFly.Visible = True
        imgDuckDie.Visible = False
        imgHit.Visible = False
        imgDogDie.Visible = False
    ElseIf imgDuckFly.Visible = True Then
        imgDuckFly.Visible = False
        imgDuckDie.Visible = False
        imgDogDie.Visible = False
        imgDogLaugh.Visible = True
        imgHaha.Visible = True
        intScore = intScore - 1
        lblScore = "Score: " & intScore
    Else
    End If
End Sub


Private Sub imgDogDie_Click()
    imgHit.Visible = False
    
    intRandomTop = Rnd * 3400
    intRandomLeft = Rnd * 6700
    
    imgDuckFly.Top = intRandomTop
    imgDuckFly.Left = intRandomLeft
    
    imgDuckFly.Visible = True
    imgDogDie.Visible = False
    imgHaha.Visible = False
End Sub

Private Sub imgDogLaugh_Click()
    imgHaha.Visible = False
    imgHit.Visible = True
    imgDogDie.Visible = True
    imgDogLaugh.Visible = False
End Sub

Private Sub imgDuckFly_Click()
    imgDogDie.Visible = False
    imgHit.Visible = True
    imgDuckFly.Visible = False
    imgDuckDie.Visible = True
    imgHaha.Visible = False

    imgDuckDie.Top = imgDuckFly.Top
    imgDuckDie.Left = imgDuckFly.Left
    
    intScore = intScore + 1
    lblScore = "Score: " & intScore
                
    Sleep 500
    
    imgHit.Visible = False
    
    intRandomTop = Rnd * 3400
    intRandomLeft = Rnd * 6700
    
    imgDuckFly.Top = intRandomTop
    imgDuckFly.Left = intRandomLeft
    
    imgDuckFly.Visible = True
    imgDuckDie.Visible = False
    imgHaha.Visible = False
End Sub

Private Sub imgHaha_Click()
    intRandomTop = Rnd * 3400
    intRandomLeft = Rnd * 6700
    
    imgDuckFly.Top = intRandomTop
    imgDuckFly.Left = intRandomLeft
    
    imgDuckFly.Visible = True
    imgDuckDie.Visible = False
    imgHaha.Visible = False
    imgDogLaugh.Visible = False
End Sub


Private Sub imgPlay_Click()
    imgDogJump.Visible = True
    imgPlay.Visible = False
    imgDogSniff.Visible = False
    
    Sleep 500
    
    imgDogJump.Visible = False
    
    intRandomTop = Rnd * 3400
    intRandomLeft = Rnd * 6700
    
    imgDuckFly.Top = intRandomTop
    imgDuckFly.Left = intRandomLeft
    
    imgDuckFly.Visible = True
    imgDuckDie.Visible = False
    imgDogDie.Visible = False
    imgDogLaugh.Visible = False
    imgHaha.Visible = False
End Sub

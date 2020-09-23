VERSION 5.00
Begin VB.Form frmDuck 
   Caption         =   "Duck Hunt-Stuart the Great"
   ClientHeight    =   6525
   ClientLeft      =   2025
   ClientTop       =   630
   ClientWidth     =   7680
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6525
   ScaleWidth      =   7680
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   360
      Top             =   0
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   0
      Top             =   0
   End
   Begin VB.OptionButton optAnthrax 
      BackColor       =   &H00000000&
      Caption         =   "Anthrax"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   1800
      MaskColor       =   &H00000000&
      MousePointer    =   2  'Cross
      TabIndex        =   5
      Top             =   6180
      Width           =   1035
   End
   Begin VB.OptionButton optMissle 
      BackColor       =   &H00000000&
      Caption         =   "Missle"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   880
      MaskColor       =   &H00000000&
      MousePointer    =   2  'Cross
      TabIndex        =   2
      Top             =   6180
      Width           =   998
   End
   Begin VB.OptionButton optShotGun 
      BackColor       =   &H80000012&
      Caption         =   "Gun"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   120
      MousePointer    =   2  'Cross
      TabIndex        =   0
      Top             =   6180
      Width           =   735
   End
   Begin VB.Label lblRound 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   3240
      TabIndex        =   6
      Top             =   2400
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Image imgRound 
      Height          =   750
      Left            =   3120
      Picture         =   "Form1.frx":030A
      Top             =   2040
      Visible         =   0   'False
      Width           =   1425
   End
   Begin VB.Image DuckDie 
      Height          =   1545
      Index           =   2
      Left            =   3720
      MousePointer    =   2  'Cross
      Picture         =   "Form1.frx":0A76
      Top             =   0
      Visible         =   0   'False
      Width           =   1800
   End
   Begin VB.Image DuckDie 
      Height          =   1395
      Index           =   1
      Left            =   2400
      MousePointer    =   2  'Cross
      Picture         =   "Form1.frx":1935
      Top             =   0
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.Image DuckDie 
      Height          =   975
      Index           =   0
      Left            =   1560
      MousePointer    =   2  'Cross
      Picture         =   "Form1.frx":21F3
      Top             =   0
      Visible         =   0   'False
      Width           =   930
   End
   Begin VB.Image imgDuckFly 
      Height          =   615
      Left            =   5520
      MousePointer    =   2  'Cross
      Picture         =   "Form1.frx":2C1E
      Top             =   360
      Visible         =   0   'False
      Width           =   810
   End
   Begin VB.Image Bullets 
      Height          =   210
      Index           =   3
      Left            =   6480
      MousePointer    =   2  'Cross
      Picture         =   "Form1.frx":32B8
      Top             =   240
      Width           =   930
   End
   Begin VB.Image Bullets 
      Height          =   210
      Index           =   2
      Left            =   6480
      Picture         =   "Form1.frx":366A
      Top             =   240
      Width           =   930
   End
   Begin VB.Image Bullets 
      Height          =   210
      Index           =   1
      Left            =   6480
      Picture         =   "Form1.frx":39D5
      Top             =   240
      Width           =   930
   End
   Begin VB.Image Bullets 
      Height          =   210
      Index           =   0
      Left            =   6480
      Picture         =   "Form1.frx":3CF3
      Top             =   240
      Width           =   930
   End
   Begin VB.Label lblScore2 
      BackStyle       =   0  'Transparent
      Caption         =   "SCORE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   6360
      MousePointer    =   2  'Cross
      TabIndex        =   4
      Top             =   5910
      Width           =   855
   End
   Begin VB.Label lblScore 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   6360
      MousePointer    =   2  'Cross
      TabIndex        =   1
      Top             =   6160
      Width           =   1215
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Height          =   615
      Index           =   1
      Left            =   6320
      Shape           =   4  'Rounded Rectangle
      Top             =   5880
      Width           =   1335
   End
   Begin VB.Image imgRed 
      Height          =   180
      Index           =   12
      Left            =   5850
      MousePointer    =   2  'Cross
      Picture         =   "Form1.frx":3FB2
      Top             =   6090
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Image imgRed 
      Height          =   180
      Index           =   11
      Left            =   5610
      MousePointer    =   2  'Cross
      Picture         =   "Form1.frx":429B
      Top             =   6090
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Image imgRed 
      Height          =   180
      Index           =   10
      Left            =   5370
      MousePointer    =   2  'Cross
      Picture         =   "Form1.frx":4584
      Top             =   6090
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Image imgRed 
      Height          =   180
      Index           =   9
      Left            =   5130
      MousePointer    =   2  'Cross
      Picture         =   "Form1.frx":486D
      Top             =   6090
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Image imgRed 
      Height          =   180
      Index           =   8
      Left            =   4890
      MousePointer    =   2  'Cross
      Picture         =   "Form1.frx":4B56
      Top             =   6090
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Image imgRed 
      Height          =   180
      Index           =   7
      Left            =   4650
      MousePointer    =   2  'Cross
      Picture         =   "Form1.frx":4E3F
      Top             =   6090
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Image imgRed 
      Height          =   180
      Index           =   6
      Left            =   4410
      MousePointer    =   2  'Cross
      Picture         =   "Form1.frx":5128
      Top             =   6090
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Image imgRed 
      Height          =   180
      Index           =   5
      Left            =   4170
      MousePointer    =   2  'Cross
      Picture         =   "Form1.frx":5411
      Top             =   6090
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Image imgRed 
      Height          =   180
      Index           =   4
      Left            =   3930
      MousePointer    =   2  'Cross
      Picture         =   "Form1.frx":56FA
      Top             =   6090
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Image imgRed 
      Height          =   180
      Index           =   3
      Left            =   3690
      MousePointer    =   2  'Cross
      Picture         =   "Form1.frx":59E3
      Top             =   6090
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Image imgRed 
      Height          =   180
      Index           =   2
      Left            =   3450
      MousePointer    =   2  'Cross
      Picture         =   "Form1.frx":5CCC
      Top             =   6090
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Image imgRed 
      Appearance      =   0  'Flat
      Height          =   180
      Index           =   1
      Left            =   3210
      MousePointer    =   2  'Cross
      Picture         =   "Form1.frx":5FB5
      Top             =   6090
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Image Image17 
      Height          =   375
      Left            =   3120
      MousePointer    =   2  'Cross
      Picture         =   "Form1.frx":629E
      Top             =   6000
      Width           =   3060
   End
   Begin VB.Label lblWeapons 
      BackColor       =   &H80000008&
      BackStyle       =   0  'Transparent
      Caption         =   "WEAPONS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   75
      MousePointer    =   2  'Cross
      TabIndex        =   3
      Top             =   5910
      Width           =   1215
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Height          =   615
      Index           =   0
      Left            =   25
      Shape           =   4  'Rounded Rectangle
      Top             =   5880
      Width           =   2895
   End
   Begin VB.Image imgBackGrass 
      Height          =   2535
      Left            =   0
      MousePointer    =   2  'Cross
      Picture         =   "Form1.frx":6A18
      Top             =   4200
      Width           =   7680
   End
   Begin VB.Image DogLaugh 
      Height          =   885
      Left            =   3360
      MousePointer    =   2  'Cross
      Picture         =   "Form1.frx":BF38
      Top             =   3480
      Visible         =   0   'False
      Width           =   1080
   End
   Begin VB.Image DogDie 
      Height          =   1065
      Index           =   0
      Left            =   3120
      MousePointer    =   2  'Cross
      Picture         =   "Form1.frx":C7D8
      Top             =   3360
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.Image DogDie 
      Height          =   1725
      Index           =   1
      Left            =   2760
      Picture         =   "Form1.frx":D196
      Top             =   3000
      Visible         =   0   'False
      Width           =   2325
   End
   Begin VB.Image DogDie 
      Height          =   1545
      Index           =   2
      Left            =   3120
      Picture         =   "Form1.frx":F03F
      Top             =   2880
      Visible         =   0   'False
      Width           =   1665
   End
   Begin VB.Image Back 
      Height          =   4455
      Left            =   0
      MousePointer    =   2  'Cross
      Picture         =   "Form1.frx":10522
      Top             =   0
      Width           =   7680
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuNewGame 
         Caption         =   "New Game"
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnuSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      Begin VB.Menu mnuHelp2 
         Caption         =   "Help"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuTip 
         Caption         =   "Tips"
      End
      Begin VB.Menu mnuSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "frmDuck"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private DX As New DirectX7
Private DSOUND As DirectSound
Private SoundGun As DirectSoundBuffer
Private SoundBuffer As DirectSoundBuffer
Private SoundBuffer02 As DirectSoundBuffer
Private Theme As DirectSoundBuffer
Private SoundDesc As DSBUFFERDESC
Private WavFormat As WAVEFORMATEX
Dim Bullet As Integer
Dim BulletTemp As Integer
Dim Weapon As Integer
Dim DuckRed As Integer
Dim Round As Integer

Private Sub Back_Click()
    BackClick
End Sub

Private Sub DogLaugh_Click()
    SoundGun.SetCurrentPosition 0
    SoundGun.Play DSBPLAY_DEFAULT
    DogLaugh.Visible = False
    SoundBuffer.Stop
    DogDie(Weapon).Visible = True
End Sub

Private Sub Form_Load()
    While BulletTemp <= 3
        Bullets(BulletTemp).Visible = False
        BulletTemp = BulletTemp + 1
    Wend
    DuckRed = 0
    Round = 1
    BulletTemp = 0
    Bullet = 3
    Bullets(Bullet).Visible = True
    Set DSOUND = DX.DirectSoundCreate("")
    DSOUND.SetCooperativeLevel hWnd, DSSCL_NORMAL
    Set SoundGun = DSOUND.CreateSoundBufferFromFile(App.Path & "\GunShot.wav", SoundDesc, WavFormat)
    Set SoundBuffer = DSOUND.CreateSoundBufferFromFile(App.Path & "\doglaugh.wav", SoundDesc, WavFormat)
    Set SoundBuffer02 = DSOUND.CreateSoundBufferFromFile(App.Path & "\DuckFly.wav", SoundDesc, WavFormat)
    Set Theme = DSOUND.CreateSoundBufferFromFile(App.Path & "\Theme.wav", SoundDesc, WavFormat)
End Sub

Private Sub BackClick()
    If DuckDie(Weapon).Visible = True Then
        DuckDie(Weapon).Visible = False
        SoundBuffer.Stop
        Bullet = 3
        While BulletTemp <= 3
            Bullets(BulletTemp).Visible = False
            BulletTemp = BulletTemp + 1
        Wend
        BulletTemp = 0
        Bullets(Bullet).Visible = True
        Timer1.Enabled = True
        Timer2.Enabled = True
    ElseIf imgRound.Visible = True Then
        imgRound.Visible = False
        lblRound.Visible = False
        Timer1.Enabled = True
        Timer2.Enabled = True
        SoundBuffer.SetCurrentPosition 0
    ElseIf DogDie(Weapon).Visible = True Or DogLaugh.Visible = True Then
        DogDie(Weapon).Visible = False
        DogLaugh.Visible = False
        SoundBuffer.Stop
        Bullet = 3
        While BulletTemp <= 3
            Bullets(BulletTemp).Visible = False
            BulletTemp = BulletTemp + 1
        Wend
        BulletTemp = 0
        Bullets(Bullet).Visible = True
        Timer1.Enabled = True
        Timer2.Enabled = True
    Else
        SoundGun.SetCurrentPosition 0
        SoundGun.Play DSBPLAY_DEFAULT
        Bullet = Bullet - 1
        If Bullet = 0 Then
            While BulletTemp <= 3
                Bullets(BulletTemp).Visible = False
                BulletTemp = BulletTemp + 1
            Wend
            BulletTemp = 0
            Bullets(Bullet).Visible = True
            DuckRed = DuckRed + 1
            imgDuckFly.Visible = False
            Timer1.Enabled = False
            SoundBuffer02.Stop
            Dog_Laugh
        ElseIf Bullet > 0 Then
            While BulletTemp <= 3
                Bullets(BulletTemp).Visible = False
                BulletTemp = BulletTemp + 1
            Wend
            BulletTemp = 0
            Bullets(Bullet).Visible = True
        Else
        End If
    End If
End Sub



Private Sub imgDuckFly_Click()
    SoundGun.SetCurrentPosition 0
    SoundGun.Play DSBPLAY_DEFAULT
    DuckRed = DuckRed + 1
    If DuckRed > 12 Then
        DuckRed = 1
        While DuckRed <= 12
            imgRed(DuckRed).Visible = False
            DuckRed = DuckRed + 1
        Wend
        DuckRed = 0
        Round = Round + 1
        Bullets(Bullet).Visible = True
        imgRound.Visible = True
        lblRound.Visible = True
        lblRound.Caption = Round
        DuckDie(Weapon).Visible = False
    Else
        imgRed(DuckRed).Visible = True
    End If
    imgDuckFly.Visible = False
    Timer1.Enabled = False
    SoundBuffer02.Stop
    DuckDie(Weapon).Top = imgDuckFly.Top
    DuckDie(Weapon).Left = imgDuckFly.Left
    DuckDie(Weapon).Visible = True
End Sub

Private Sub mnuAbout_Click()
    frmAbout.Show
End Sub

Private Sub mnuExit_Click()
    Unload Me
End Sub

Private Sub mnuHelp2_Click()
    frmHelp.Show
End Sub

Private Sub mnuNewGame_Click()
    DuckDie(Weapon).Visible = False
    DogDie(Weapon).Visible = False
    DogLaugh.Visible = False
    SoundBuffer.Stop
    SoundBuffer02.Stop
    SoundGun.Stop
    Bullet = 3
    While BulletTemp <= 3
        Bullets(BulletTemp).Visible = False
        BulletTemp = BulletTemp + 1
    Wend
    BulletTemp = 0
    DuckRed = 1
    While DuckRed <= 12
        imgRed(DuckRed).Visible = False
        DuckRed = DuckRed + 1
    Wend
    DuckRed = 0
    Round = 1
    Bullets(Bullet).Visible = True
    imgRound.Visible = True
    lblRound.Visible = True
    lblRound.Caption = Round
End Sub

Private Sub mnuTip_Click()
    frmTip.Show
End Sub

Private Sub optAnthrax_Click()
    Weapon = 2
End Sub

Private Sub optMissle_Click()
    Weapon = 1
    Set SoundGun = DSOUND.CreateSoundBufferFromFile(App.Path & "\GunShot.wav", SoundDesc, WavFormat)
End Sub

Private Sub optShotGun_Click()
    Weapon = 0
    Set SoundGun = DSOUND.CreateSoundBufferFromFile(App.Path & "\GunShot.wav", SoundDesc, WavFormat)
End Sub

Private Sub Timer1_Timer()
    SoundBuffer02.Play DSBPLAY_LOOP
    imgDuckFly.Visible = True
    If imgDuckFly.Visible = True Then
        If Timer1.Interval = 1 Then
            imgDuckFly.Left = imgDuckFly.Left - 50
            If imgDuckFly.Left < 0 Then
                Timer1.Interval = 2
            Else
                imgDuckFly.Left = imgDuckFly.Left - 50
            End If
        End If
    
        If Timer1.Interval = 2 Then
            imgDuckFly.Left = imgDuckFly.Left + 50
            If imgDuckFly.Left > frmDuck.Width Then
                Timer1.Interval = 1
            Else
                imgDuckFly.Left = imgDuckFly.Left + 50
            End If
        End If
    End If
End Sub

Private Sub Timer2_Timer()
    If imgDuckFly.Visible = True Then
        If Timer2.Interval = 1 Then
            imgDuckFly.Top = imgDuckFly.Top + 50
            If imgDuckFly.Top > 4200 Then
                Timer2.Interval = 2
            Else
                imgDuckFly.Top = imgDuckFly.Top + 50
            End If
        End If
    
        If Timer2.Interval = 2 Then
            imgDuckFly.Top = imgDuckFly.Top - 50
            If imgDuckFly.Top < 0 Then
                Timer2.Interval = 1
            Else
                imgDuckFly.Top = imgDuckFly.Top - 50
            End If
        End If
    End If
End Sub

Private Sub Dog_Laugh()
    DogLaugh.Visible = True
    SoundBuffer.SetCurrentPosition 0
    SoundBuffer.Play DSBPLAY_DEFAULT
End Sub


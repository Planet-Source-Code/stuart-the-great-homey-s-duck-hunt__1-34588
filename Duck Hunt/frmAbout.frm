VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About-Stuart the Great"
   ClientHeight    =   2175
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   5610
   ClipControls    =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1501.224
   ScaleMode       =   0  'User
   ScaleWidth      =   5268.08
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox picIcon 
      AutoSize        =   -1  'True
      ClipControls    =   0   'False
      Height          =   1035
      Left            =   120
      Picture         =   "frmAbout.frx":0000
      ScaleHeight     =   684.775
      ScaleMode       =   0  'User
      ScaleWidth      =   653.17
      TabIndex        =   1
      Top             =   120
      Width           =   990
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   4320
      TabIndex        =   0
      Top             =   1800
      Width           =   1260
   End
   Begin VB.Label Label1 
      Caption         =   "Website: http://www.HomeysCrib.tk"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1200
      TabIndex        =   6
      Top             =   1200
      Width           =   4455
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   0
      X2              =   5224.884
      Y1              =   1066.387
      Y2              =   1066.387
   End
   Begin VB.Label lblDescription 
      Caption         =   "Description: Shoot all the ducks..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   210
      Left            =   1200
      TabIndex        =   2
      Top             =   840
      Width           =   4365
   End
   Begin VB.Label lblTitle 
      Caption         =   "Homey's Duck Hunt"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Left            =   1200
      TabIndex        =   4
      Top             =   120
      Width           =   3885
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   0
      X2              =   5210.798
      Y1              =   1076.74
      Y2              =   1076.74
   End
   Begin VB.Label lblVersion 
      Caption         =   "Version 1.00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   1200
      TabIndex        =   5
      Top             =   480
      Width           =   3885
   End
   Begin VB.Label lblDisclaimer 
      Caption         =   "WARNING: Shooting real ducks is not nice to ducks, same with shooting real dogs..."
      ForeColor       =   &H00000000&
      Height          =   465
      Left            =   120
      TabIndex        =   3
      Top             =   1680
      Width           =   4095
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOK_Click()
  Unload Me
End Sub


VERSION 5.00
Begin VB.Form frmTip 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tip of the Day-Stuart the Great"
   ClientHeight    =   1215
   ClientLeft      =   2355
   ClientTop       =   2385
   ClientWidth     =   5415
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1215
   ScaleWidth      =   5415
   StartUpPosition =   1  'CenterOwner
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin VB.CommandButton cmdNextTip 
      Caption         =   "&Next Tip"
      Height          =   375
      Left            =   4080
      TabIndex        =   2
      Top             =   720
      Width           =   1215
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   1035
      Left            =   120
      Picture         =   "frmTip.frx":0000
      ScaleHeight     =   975
      ScaleWidth      =   3675
      TabIndex        =   1
      Top             =   120
      Width           =   3735
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Did you know..."
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
         Left            =   480
         TabIndex        =   4
         Top             =   120
         Width           =   2655
      End
      Begin VB.Label lblTipText 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "You can click on the ducks."
         Height          =   315
         Left            =   120
         TabIndex        =   3
         Top             =   600
         Width           =   3375
      End
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   4080
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmTip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim intTips As Integer

Private Sub cmdNextTip_Click()
    If intTips = 0 Then
        intTips = intTips + 1
        lblTipText.Caption = "That's all..."
    ElseIf intTips = 1 Then
        intTips = intTips + 1
        lblTipText.Caption = "Stop clicking that, there's no more."
    ElseIf intTips = 2 Then
        intTips = intTips + 1
        lblTipText.Caption = "STOP!"
    ElseIf intTips = 3 Then
        intTips = 0
        lblTipText.Caption = "NO MORE TIPS!!!"
    End If
End Sub

Private Sub cmdOK_Click()
    Unload frmTip
End Sub

Private Sub Form_Load()
    intTips = 0
End Sub

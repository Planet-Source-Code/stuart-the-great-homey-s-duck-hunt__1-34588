VERSION 5.00
Begin VB.Form frmSplash 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4245
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   7380
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   2  'Cross
   ScaleHeight     =   4245
   ScaleWidth      =   7380
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   4050
      Left            =   150
      TabIndex        =   0
      Top             =   60
      Width           =   7080
      Begin VB.Label lblWarning 
         BackStyle       =   0  'Transparent
         Caption         =   " Warning: this game is boring."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   120
         MousePointer    =   2  'Cross
         TabIndex        =   1
         Top             =   3720
         Width           =   6975
      End
      Begin VB.Image Image1 
         Height          =   3900
         Left            =   0
         MousePointer    =   2  'Cross
         Picture         =   "frmSplash.frx":000C
         Top             =   120
         Width           =   7095
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00000000&
         BackStyle       =   1  'Opaque
         Height          =   275
         Left            =   0
         Top             =   100
         Width           =   7095
      End
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Frame1_Click()
    Unload Me
    frmDuck.Show
End Sub

Private Sub Image1_Click()
    Unload Me
    frmDuck.Show
End Sub

Private Sub lblWarning_Click()
    Unload Me
    frmDuck.Show
End Sub

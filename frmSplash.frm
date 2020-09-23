VERSION 5.00
Begin VB.Form frmSplash 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2385
   ClientLeft      =   225
   ClientTop       =   1380
   ClientWidth     =   3900
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2385
   ScaleWidth      =   3900
   StartUpPosition =   2  'CenterScreen
   Begin VB.Image Image5 
      Height          =   480
      Left            =   3000
      Picture         =   "frmSplash.frx":000C
      Top             =   480
      Width           =   480
   End
   Begin VB.Image Image4 
      Height          =   480
      Left            =   720
      Picture         =   "frmSplash.frx":08D6
      Top             =   1560
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   3000
      Picture         =   "frmSplash.frx":15A0
      Top             =   1560
      Width           =   480
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Patient Information System"
      BeginProperty Font 
         Name            =   "Lucida Handwriting"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1095
      Left            =   1320
      TabIndex        =   0
      Top             =   600
      Width           =   3135
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Patient Information System"
      BeginProperty Font 
         Name            =   "Lucida Handwriting"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1095
      Left            =   1330
      TabIndex        =   1
      Top             =   610
      Width           =   3135
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   720
      Picture         =   "frmSplash.frx":226A
      Top             =   480
      Width           =   480
   End
   Begin VB.Image Image3 
      Height          =   2400
      Left            =   -120
      Picture         =   "frmSplash.frx":2F34
      Stretch         =   -1  'True
      Top             =   0
      Width           =   4440
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_KeyPress(KeyAscii As Integer)
Unload Me
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Unload Me
End Sub

Private Sub Image3_Click()
Unload Me
End Sub

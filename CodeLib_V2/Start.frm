VERSION 5.00
Begin VB.Form Start 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4680
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7470
   LinkTopic       =   "Form1"
   Moveable        =   0   'False
   ScaleHeight     =   312
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   498
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Image Image4 
      Height          =   1005
      Left            =   360
      Picture         =   "Start.frx":0000
      Top             =   405
      Width           =   6735
   End
   Begin VB.Image Image3 
      Height          =   1185
      Left            =   1080
      Picture         =   "Start.frx":2050
      Top             =   1575
      Width           =   5250
   End
   Begin VB.Image Image1 
      Height          =   1920
      Left            =   45
      Picture         =   "Start.frx":4949
      Top             =   2745
      Width           =   3165
   End
   Begin VB.Image Image2 
      Height          =   1920
      Left            =   4230
      Picture         =   "Start.frx":74DA
      Top             =   2745
      Width           =   3165
   End
End
Attribute VB_Name = "Start"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Click()
Start.Hide
End Sub

Private Sub Form_Load()
Call ColBar(Start, 190, 90, 40, 4, 0, 128, 128, 64)
Call ColBox(Start, 0, 0, Start.ScaleWidth, Start.ScaleHeight, 21, 64, 64, 0, 198, 190, 128)
Start.Move (Screen.Width / 2) - (Start.Width / 2), (Screen.Height / 2) - (Start.Height / 2)
ExplodeForm Start, 30, 0
End Sub

Private Sub Image1_Click()
Start.Hide
End Sub

Private Sub Image2_Click()
Start.Hide
End Sub

Private Sub Image3_Click()
Start.Hide
End Sub

Private Sub Image4_Click()
Start.Hide
End Sub

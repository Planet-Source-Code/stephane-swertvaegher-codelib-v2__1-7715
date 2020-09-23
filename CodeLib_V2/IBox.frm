VERSION 5.00
Begin VB.Form IBox 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2175
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4890
   FillColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   Moveable        =   0   'False
   ScaleHeight     =   145
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   326
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0C000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   180
      TabIndex        =   6
      Text            =   "Text1"
      Top             =   1710
      Width           =   3210
   End
   Begin VB.PictureBox But1 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   375
      Index           =   1
      Left            =   3510
      Picture         =   "IBox.frx":0000
      ScaleHeight     =   25
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   84
      TabIndex        =   3
      Top             =   1620
      Width           =   1260
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Exit"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFC0&
         Height          =   240
         Index           =   1
         Left            =   45
         TabIndex        =   5
         Top             =   90
         Width           =   1170
      End
   End
   Begin VB.PictureBox But1 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   375
      Index           =   0
      Left            =   3510
      Picture         =   "IBox.frx":03C8
      ScaleHeight     =   25
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   84
      TabIndex        =   2
      Top             =   1170
      Width           =   1260
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Accept"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFC0&
         Height          =   240
         Index           =   0
         Left            =   45
         TabIndex        =   4
         Top             =   90
         Width           =   1170
      End
   End
   Begin VB.Image Image4 
      Height          =   480
      Left            =   3960
      Picture         =   "IBox.frx":0790
      Top             =   540
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   4275
      Top             =   630
      Width           =   480
   End
   Begin VB.Image Image3 
      Height          =   375
      Left            =   2745
      Picture         =   "IBox.frx":1352
      Top             =   1530
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Image Image2 
      Height          =   375
      Left            =   2790
      Picture         =   "IBox.frx":171C
      Top             =   1050
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0C0&
      Height          =   975
      Left            =   180
      TabIndex        =   1
      Top             =   495
      Width           =   3015
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "This is the title of the inputbox"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFC0&
      Height          =   195
      Left            =   180
      TabIndex        =   0
      Top             =   90
      Width           =   4500
   End
End
Attribute VB_Name = "IBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub But1_GotFocus(Index As Integer)
Text1.SetFocus
End Sub

Private Sub Form_Activate()
On Error Resume Next
Text1.SetFocus
End Sub

Private Sub But1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
But1(Index).Picture = Image3.Picture
Label3(Index).Left = 5
Label3(Index).Top = 8
End Sub

Private Sub But1_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
But1(Index).Picture = Image2.Picture
Label3(Index).Left = 3
Label3(Index).Top = 6
IbReturn = ""
If Index = 0 Then IbReturn = Text1.Text
Unload Me
End Sub

Private Sub Label3_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
But1(Index).Picture = Image3.Picture
Label3(Index).Left = Label3(Index).Left + 1
Label3(Index).Top = Label3(Index).Top + 1
End Sub

Private Sub Label3_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
But1(Index).Picture = Image2.Picture
Label3(Index).Left = Label3(Index).Left - 1
Label3(Index).Top = Label3(Index).Top - 1
IbReturn = ""
If Index = 0 Then IbReturn = Text1.Text
Unload Me
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    IbReturn = Text1.Text
Unload Me
End If
End Sub

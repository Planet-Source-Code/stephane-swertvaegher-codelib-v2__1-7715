VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form MBox 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   1800
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4005
   FillColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   Moveable        =   0   'False
   ScaleHeight     =   120
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   267
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   270
      Top             =   675
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MBox.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MBox.frx":0E54
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MBox.frx":1CA8
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MBox.frx":2AFC
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MBox.frx":3950
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MBox.frx":565C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MBox.frx":64B0
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MBox.frx":7084
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MBox.frx":7960
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox But1 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   375
      Index           =   1
      Left            =   1845
      Picture         =   "MBox.frx":8534
      ScaleHeight     =   25
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   84
      TabIndex        =   3
      Top             =   1305
      Width           =   1260
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Color"
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
      Left            =   225
      Picture         =   "MBox.frx":88FC
      ScaleHeight     =   25
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   84
      TabIndex        =   2
      Top             =   1305
      Width           =   1260
      Begin VB.Label Label3 
         Alignment       =   2  'Center
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
         ForeColor       =   &H00FFFFC0&
         Height          =   240
         Index           =   0
         Left            =   45
         TabIndex        =   4
         Top             =   90
         Width           =   1170
      End
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   3420
      Top             =   630
      Width           =   480
   End
   Begin VB.Image Image3 
      Height          =   375
      Left            =   2745
      Picture         =   "MBox.frx":8CC4
      Top             =   855
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Image Image2 
      Height          =   375
      Left            =   2790
      Picture         =   "MBox.frx":908E
      Top             =   1350
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   690
      Left            =   180
      TabIndex        =   1
      Top             =   495
      Width           =   3015
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "This is the title of the messagebox"
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
      Left            =   495
      TabIndex        =   0
      Top             =   90
      Width           =   2925
   End
End
Attribute VB_Name = "MBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub But1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
But1(Index).Picture = Image3.Picture
Label3(Index).Left = 5 'Label3(Index).Left + 2
Label3(Index).Top = 8 'Label3(Index).Top + 2
End Sub

Private Sub But1_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
But1(Index).Picture = Image2.Picture
Label3(Index).Left = 3 'Label3(Index).Left - 2
Label3(Index).Top = 6 'Label3(Index).Top - 2
MBReturn = Index
Unload Me
End Sub

Private Sub Label3_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
But1(Index).Picture = Image3.Picture
Label3(Index).Left = 5 'Label3(Index).Left + 2
Label3(Index).Top = 8 'Label3(Index).Top + 2
End Sub

Private Sub Label3_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
But1(Index).Picture = Image2.Picture
Label3(Index).Left = 3 'Label3(Index).Left - 2
Label3(Index).Top = 6 'Label3(Index).Top - 2
MBReturn = Index
Unload Me
End Sub


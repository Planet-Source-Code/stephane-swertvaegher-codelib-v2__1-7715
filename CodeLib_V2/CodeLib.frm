VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form CodeLib 
   AutoRedraw      =   -1  'True
   Caption         =   "CodeLib V2.0"
   ClientHeight    =   6660
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10470
   Icon            =   "CodeLib.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   444
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   698
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Pic5 
      BorderStyle     =   0  'None
      Height          =   195
      Left            =   7065
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   19
      Top             =   1305
      Width           =   195
   End
   Begin VB.PictureBox Pic1 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   9765
      Picture         =   "CodeLib.frx":030A
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   8
      Top             =   1170
      Width           =   480
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Height          =   780
      Left            =   90
      TabIndex        =   5
      Top             =   180
      Width           =   10245
      _ExtentX        =   18071
      _ExtentY        =   1376
      ButtonWidth     =   1561
      ButtonHeight    =   1376
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   15
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Save DB"
            Key             =   "key1"
            Object.ToolTipText     =   "Save the Database"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Print Text"
            Key             =   "key2"
            Object.ToolTipText     =   "Print Code or Helpfile"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Copy Code"
            Key             =   "key3"
            Object.ToolTipText     =   "Copy selected Code to Clipboard"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Kill Code"
            Key             =   "key4"
            Object.ToolTipText     =   "Remove selected Code/Help/Notes"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Rename"
            Key             =   "key7"
            Object.ToolTipText     =   "Rename selected Code"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Move"
            Key             =   "key10"
            Object.ToolTipText     =   "Move selected Code to another Category"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Category"
            Key             =   "key5"
            Object.ToolTipText     =   "Add new Category"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Rename"
            Key             =   "key6"
            Object.ToolTipText     =   "Rename selected Category"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Kill Categ."
            Key             =   "key11"
            Object.ToolTipText     =   "Remove selected Category"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Search"
            Key             =   "key8"
            Object.ToolTipText     =   "Search for Code or Helpfiles"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Help"
            Key             =   "key9"
            Object.ToolTipText     =   "Help-Files"
            ImageIndex      =   9
         EndProperty
      EndProperty
   End
   Begin VB.ComboBox Combo1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   90
      TabIndex        =   4
      Text            =   "Combo1"
      Top             =   1710
      Width           =   2400
   End
   Begin VB.ListBox List1 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H00FF0000&
      Height          =   3180
      Left            =   90
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   2835
      Width           =   2400
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4965
      Left            =   2610
      TabIndex        =   0
      Top             =   1710
      Width           =   7845
      _ExtentX        =   13838
      _ExtentY        =   8758
      _Version        =   393216
      Tabs            =   4
      Tab             =   2
      TabsPerRow      =   4
      TabHeight       =   476
      ShowFocusRect   =   0   'False
      BackColor       =   12632256
      ForeColor       =   16711680
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Code"
      TabPicture(0)   =   "CodeLib.frx":0BD4
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Pic7"
      Tab(0).Control(1)=   "Pic8"
      Tab(0).Control(2)=   "ImageList1"
      Tab(0).Control(3)=   "Text1(0)"
      Tab(0).Control(4)=   "Label10"
      Tab(0).Control(5)=   "Label7"
      Tab(0).Control(6)=   "Label8"
      Tab(0).Control(7)=   "Label1(0)"
      Tab(0).ControlCount=   8
      TabCaption(1)   =   "Help"
      TabPicture(1)   =   "CodeLib.frx":0BF0
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Pic2"
      Tab(1).Control(1)=   "Text1(1)"
      Tab(1).Control(2)=   "Label1(1)"
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "Notes"
      TabPicture(2)   =   "CodeLib.frx":0C0C
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "Label1(2)"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Text1(2)"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Pic3"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).ControlCount=   3
      TabCaption(3)   =   "Edit Code"
      TabPicture(3)   =   "CodeLib.frx":0C28
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Label1(3)"
      Tab(3).Control(1)=   "Text1(3)"
      Tab(3).Control(2)=   "Pic4"
      Tab(3).ControlCount=   3
      Begin VB.PictureBox Pic7 
         BackColor       =   &H00FF00FF&
         Height          =   195
         Left            =   -70725
         ScaleHeight     =   135
         ScaleWidth      =   135
         TabIndex        =   22
         Top             =   360
         Width           =   195
      End
      Begin VB.PictureBox Pic8 
         BackColor       =   &H00FF00FF&
         Height          =   195
         Left            =   -70725
         ScaleHeight     =   135
         ScaleWidth      =   135
         TabIndex        =   21
         Top             =   585
         Width           =   195
      End
      Begin VB.PictureBox Pic4 
         Height          =   195
         Left            =   -70050
         ScaleHeight     =   135
         ScaleWidth      =   135
         TabIndex        =   17
         Top             =   540
         Width           =   195
      End
      Begin VB.PictureBox Pic3 
         Height          =   195
         Left            =   4275
         ScaleHeight     =   135
         ScaleWidth      =   135
         TabIndex        =   15
         Top             =   540
         Width           =   195
      End
      Begin VB.PictureBox Pic2 
         Height          =   195
         Left            =   -70725
         ScaleHeight     =   9
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   9
         TabIndex        =   14
         Top             =   540
         Width           =   195
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00A0D0D0&
         ForeColor       =   &H00FF0000&
         Height          =   4065
         Index           =   3
         Left            =   -74865
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   13
         Top             =   810
         Width           =   7575
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00A0D0D0&
         ForeColor       =   &H00FF0000&
         Height          =   4065
         Index           =   2
         Left            =   135
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   10
         Top             =   810
         Width           =   7575
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00A0D0D0&
         ForeColor       =   &H00FF0000&
         Height          =   4065
         Index           =   1
         Left            =   -74865
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   9
         Top             =   810
         Width           =   7575
      End
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   -74415
         Top             =   3348
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   32
         ImageHeight     =   32
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   11
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "CodeLib.frx":0C44
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "CodeLib.frx":1520
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "CodeLib.frx":20F4
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "CodeLib.frx":2F48
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "CodeLib.frx":3D9C
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "CodeLib.frx":4678
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "CodeLib.frx":4F54
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "CodeLib.frx":5830
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "CodeLib.frx":610C
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "CodeLib.frx":7E18
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "CodeLib.frx":9B24
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00A0D0D0&
         ForeColor       =   &H00FF0000&
         Height          =   4065
         Index           =   0
         Left            =   -74865
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   2
         Top             =   810
         Width           =   7575
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
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
         Height          =   255
         Left            =   -68970
         TabIndex        =   27
         Top             =   495
         Width           =   1680
      End
      Begin VB.Label Label7 
         Caption         =   "Helpfile present"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   240
         Left            =   -70455
         TabIndex        =   24
         Top             =   360
         Width           =   1500
      End
      Begin VB.Label Label8 
         Caption         =   "Notes present"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   240
         Left            =   -70455
         TabIndex        =   23
         Top             =   585
         Width           =   1500
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C000&
         BorderStyle     =   1  'Fixed Single
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
         Height          =   255
         Index           =   3
         Left            =   -74865
         TabIndex        =   16
         Top             =   495
         Width           =   4785
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C000&
         BorderStyle     =   1  'Fixed Single
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
         Height          =   255
         Index           =   2
         Left            =   135
         TabIndex        =   12
         Top             =   495
         Width           =   4020
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C000&
         BorderStyle     =   1  'Fixed Single
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
         Height          =   255
         Index           =   1
         Left            =   -74865
         TabIndex        =   11
         Top             =   495
         Width           =   4020
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C000&
         BorderStyle     =   1  'Fixed Single
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
         Height          =   255
         Index           =   0
         Left            =   -74865
         TabIndex        =   3
         Top             =   495
         Width           =   4020
      End
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackColor       =   &H00C00000&
      Caption         =   " Number of Code-snippets: 111"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF80&
      Height          =   420
      Left            =   135
      TabIndex        =   26
      Top             =   6165
      Width           =   2310
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   195
      Left            =   135
      TabIndex        =   25
      Top             =   2520
      Width           =   2265
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Caption         =   "Database dirty"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   195
      Left            =   7470
      TabIndex        =   20
      Top             =   1305
      Width           =   1800
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "Clipboard empty"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   195
      Left            =   2700
      TabIndex        =   18
      Top             =   1305
      Width           =   3750
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "Loading database..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   195
      Left            =   135
      TabIndex        =   7
      Top             =   1305
      Width           =   2355
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   195
      Left            =   135
      TabIndex        =   6
      Top             =   2160
      Width           =   2265
   End
   Begin VB.Menu mnuCode 
      Caption         =   "Code"
      Visible         =   0   'False
      Begin VB.Menu mnuCopyCode 
         Caption         =   "Copy Code"
      End
      Begin VB.Menu bar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuKillCode 
         Caption         =   "Kill Code"
      End
      Begin VB.Menu bar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRenameCode 
         Caption         =   "Rename Code"
      End
      Begin VB.Menu bar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMoveCode 
         Caption         =   "Move Code"
      End
      Begin VB.Menu bar3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPrintCod 
         Caption         =   "Print Code"
      End
   End
End
Attribute VB_Name = "CodeLib"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Dim YPos%, TH%

Private Sub Combo1_Click()
List1.Clear
For xx = 0 To 2
Label1(xx).Caption = Empty
Text1(xx).Text = ""
Next xx
Label2.Caption = Combo1.List(Combo1.ListIndex)
CatIdx = Val(Left(Combo1.List(Combo1.ListIndex), 2))
'Label3.Caption = "Category: " & CatIdx
SearchItems

Pic7.Visible = False
Label7.Visible = False
Pic8.Visible = False
Label8.Visible = False
Pic2.BackColor = RGB(192, 192, 192)
Pic3.BackColor = RGB(192, 192, 192)
Pic1.SetFocus
End Sub

Private Sub Form_Activate()
Combo1.Text = "Category"
Label4.Caption = "Ready !"
Label4.Tag = ""
If SSTab1.Tab <> 0 Then
End If
End Sub

Private Sub Form_Load()
On Error Resume Next
Title = "CodeLib V2.0"
CodeLib.ScaleMode = 1
TH = Me.TextHeight(a)
CodeLib.ScaleMode = 3
T3D CodeLib, Toolbar1, 3, T3dRaiseInset
T3D CodeLib, Label4, 5, T3dRaiseRaise
T3D CodeLib, Label5, 5, T3dRaiseRaise
T3D CodeLib, Pic5, 5, T3dRaiseRaise
T3D CodeLib, Label2, 3, T3dRaiseRaise
T3D CodeLib, Label9, 3, T3dRaiseRaise
T3D CodeLib, Label6, 5, T3dRaiseRaise
T3D Search, Search.Label1(0), 10, T3dRaiseInset
T3D Search, Search.Label1(0), 5, T3dRaiseRaise
T3D Search, Search.Drive1, 5, T3dRaiseRaise
T3D Search, Search.Dir1, 5, T3dRaiseRaise
T3D Search, Search.File1, 5, T3dRaiseRaise
T3D Search, Search.Text1, 5, T3dRaiseRaise
T3D Search2, Search2.Text1, 5, T3dRaiseRaise
T3D Search2, Search2.Combo1, 3, T3dRaiseRaise
T3D Search2, Search2.Label1, 7, T3dRaiseRaise
T3D Search2, Search2.Label2, 5, T3dRaiseRaise
T3D Search2, Search2.Label3, 5, T3dRaiseRaise
T3D Search3, Search3.Label1, 7, T3dRaiseRaise
T3D Search3, Search3.Combo1, 5, T3dRaiseRaise
T3D Move1, Move1.Combo1, 5, T3dRaiseRaise
T3D Move1, Move1.Label1, 5, T3dRaiseRaise
T3D HelpScreen, HelpScreen.Label1, 4, T3dInsetInset
T3D HelpScreen, HelpScreen.Label1, 8, T3dInsetInset
T3D HelpScreen, HelpScreen.Label1, 12, T3dInsetInset
'Label10.Caption = " Lines of text: "
'T3D CodeLib, Label10, 3, T3dRaiseRaise
T3D CodeLib, Label11, 3, T3dInsetInset
Setline CodeLib, 1
Setline CodeLib, 74, False
Pic2.BackColor = RGB(192, 192, 192)
Pic3.BackColor = RGB(192, 192, 192)
Pic4.BackColor = RGB(192, 192, 192)
Pic5.BackColor = RGB(192, 192, 192)
Label6.Caption = "Database clear"
Label4.Tag = "0"
Label4.Caption = "Loading Database"
SSTab1.Tab = 0
Pic7.Visible = False
Label7.Visible = False
Pic8.Visible = False
Label8.Visible = False
LoadCat
LoadLib
SizeCombo CodeLib, Combo1
SizeCombo Move1, Move1.Combo1
SizeCombo Search3, Search3.Combo1
Label11.Caption = "Number of Code-snippets:" & vbCr & CodeCount
Search.Drive1.Drive = "C:\"
Search.Dir1.Path = "C:\"
CodeLib.Move (Screen.Width / 2) - (CodeLib.Width / 2), (Screen.Height / 2) - (CodeLib.Height / 2), 10590, 7065
CodeLib.Show
DoEvents
Start.Show 1
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Label4.Tag = "" Then Label4.Caption = ""
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Cancel = True
Msbox "Do you want to quit ?", Title, mbYesNo, mbQuestion
If MBReturn = 1 Then Exit Sub
'OK !Quit

If Pic5.BackColor = 0 Then 'database dirty
    Msbox "The database is dirty and should be saved..." & vbCr & vbCr & "Save the Database ?", Title, mbYesNo, mbQuestion
    If MBReturn = 1 Then End
    SaveLib
End If
End
End Sub

Private Sub Form_Resize()
On Error Resume Next
CodeLib.Move (Screen.Width / 2) - (CodeLib.Width / 2), (Screen.Height / 2) - (CodeLib.Height / 2), 10590, 7065
End Sub

Private Sub List1_Click()
SSTab1.Tab = 0
Idx1 = List1.ListIndex
For xx = 0 To 2
Label1(xx).Caption = Label2.Caption & " " & List1.List(Idx1)
Text1(xx).Text = CLdata(xx + 2, List1.ItemData(Idx1))
Next xx
If CLdata(3, List1.ItemData(Idx1)) <> "" Then
Pic7.Visible = True
Label7.Visible = True
Else
Pic7.Visible = False
Label7.Visible = False
End If
If CLdata(4, List1.ItemData(Idx1)) <> "" Then
Pic8.Visible = True
Label8.Visible = True
Else
Pic8.Visible = False
Label8.Visible = False
End If
Text1(3).Text = Text1(0).Text 'copy to editbox
Label1(3).Caption = "Edit Code for: " & List1.List(Idx1)
Pic2.BackColor = RGB(192, 192, 192)
Pic3.BackColor = RGB(192, 192, 192)
Pic4.BackColor = RGB(192, 192, 192)

End Sub

Private Sub mnuCopyCode_Click()
If Text1(0).Text = "" Then
    Msbox "No Code selected", Title, mbOkonly, mbInfo
    Exit Sub
End If
Clipboard.Clear
Clipboard.SetText Text1(0).Text
Label5.Caption = "Clipboard filled with " & List1.List(Idx1)

End Sub

Private Sub mnuKillCode_Click()
If Text1(0).Text = "" Then
    Msbox "No Code selected", Title, mbOkonly, mbInfo
    Exit Sub
End If
Msbox "You want to delete the Code " & List1.List(Idx1) & " in the category " & Label2.Caption & vbCr & "Also the helpfile and notes will be deleted from the database..." & vbCr & vbCr & "Are you sure about this ?", Title, mbYesNo, mbQuestion
If MBReturn = 1 Then Exit Sub
KillEntry
End Sub

Private Sub mnuMoveCode_Click()
If Text1(0).Text = "" Then
Msbox "No Code selected...", Title, mbOkonly, mbInfo
Exit Sub
End If
Move1.Label1.Caption = "Move the selected Code:" & vbCr & List1.List(List1.ListIndex)
Move1.Show 1
End Sub

Private Sub mnuPrintCod_Click()
If Text1(0).Text = "" Then
    Msbox "No Code selected", Title, mbOkonly, mbInfo
    Exit Sub
End If
If SSTab1.Tab = 2 Or SSTab1.Tab = 3 Then
Msbox "No printing here...", Title, mbOkonly, mbInfo
Exit Sub
End If
If SSTab1.Tab = 1 And Text1(1).Text = "" Then
    Msbox "No Helpfile present...", Title, mbOkonly, mbInfo
    Exit Sub
End If
If SSTab1.Tab = 0 Then Msbox "Print Code of the " & List1.List(Idx1) & " ?", Title, mbYesNo, mbQuestion
If SSTab1.Tab = 1 Then Msbox "Print helpfile of the " & List1.List(Idx1) & " ?", Title, mbYesNo, mbQuestion
If MBReturn = 1 Then Exit Sub 'No selected
'OK ! Print !
Msbox "Turn on the printer...", Title, mbPrintNoWay, mbPrint
If MBReturn = 1 Then Exit Sub
Printer.FontSize = 10
Printer.Print
Printer.Print
Printer.CurrentX = 1000
Printer.Print Label1(SSTab1.Tab).Caption
Printer.Print
Printer.Print Text1(SSTab1.Tab).Text
Printer.EndDoc
End Sub

Private Sub mnuRenameCode_Click()
If Text1(0).Text = "" Then
Msbox "No Code selected...", Title, mbOkonly, mbInfo
Exit Sub
End If
key72:
InBox "Rename the Code: " & vbCr & CLdata(1, List1.ItemData(Idx1)), CLdata(1, List1.ItemData(Idx1)), Title
If IbReturn = "" Then Exit Sub 'exit
    For xx = 0 To 999
    If LCase(Trim(IbReturn)) = LCase(CLdata(1, xx)) Then
    Msbox "This Codename already exists !", Title, mbOkonly, mbCritical
    xx = 999
    GoTo key72
    End If
    Next xx
RenameCode
End Sub

Private Sub pic1_GotFocus()
Combo1.Text = "Category"
End Sub

Private Sub List1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    YPos = Y \ TH + List1.TopIndex
        If YPos < List1.ListCount Then
        Label4.Caption = "(" & CatIdx & ") " & List1.List(YPos)
        Else
        Label4.Caption = ""
        End If
End Sub

Private Sub Text1_LostFocus(Index As Integer)
If Index = 1 And Pic2.BackColor = 0 Then
Pic2.BackColor = RGB(192, 192, 192)
Pic5.BackColor = 0
    If CLdata(3, List1.ItemData(Idx1)) <> "" Then
    Pic7.Visible = True
    Label7.Visible = True
    Else
    Pic7.Visible = False
    Label7.Visible = False
    End If

Label6.Caption = "Database dirty"
CLdata(3, List1.ItemData(Idx1)) = Text1(1).Text
End If

If Index = 2 And Pic3.BackColor = 0 Then 'notes dirty
Pic3.BackColor = RGB(192, 192, 192)
Pic5.BackColor = 0
    If CLdata(4, List1.ItemData(Idx1)) <> "" Then
    Pic8.Visible = True
    Label8.Visible = True
    Else
    Pic8.Visible = False
    Label8.Visible = False
    End If
Label6.Caption = "Database dirty"
CLdata(4, List1.ItemData(Idx1)) = Text1(2).Text
End If

If Index = 3 And Pic4.BackColor = 0 Then 'Edit code dirty
Msbox "The Code has been changed !" & vbCr & vbCr & "Do you want to replace the actual Code with the new one ?" & vbCr & vbCr & "Note that the changed Code can not be tested !", Title, mbOKCancel, mbQuestion
Pic4.BackColor = RGB(192, 192, 192)
    If MBReturn = 1 Then 'cancel selected
    Text1(3).Text = CLdata(2, List1.ItemData(Idx1))
    Pic4.BackColor = RGB(192, 192, 192)
    Exit Sub
    End If
    'Ok ! Copy everything
    CLdata(2, List1.ItemData(Idx1)) = Text1(3).Text
    Text1(0).Text = Text1(3).Text
    Pic5.BackColor = 0
    Label6.Caption = "Database dirty"
End If

End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
T3D CodeLib, Label10, 3, T3dNone
Label10.Caption = " Lines of text: " & GetLineCount(Text1(SSTab1.Tab)) & " "
T3D CodeLib, Label10, 3, T3dRaiseRaise
Pic1.SetFocus
DoEvents
On Error Resume Next
If CLdata(3, List1.ItemData(Idx1)) <> "" Then
Pic7.Visible = True
Label7.Visible = True
Else
Pic7.Visible = False
Label7.Visible = False
End If
If CLdata(4, List1.ItemData(Idx1)) <> "" Then
Pic8.Visible = True
Label8.Visible = True
Else
Pic8.Visible = False
Label8.Visible = False
End If

If SSTab1.Tab <> 0 And Text1(0).Text = "" Then
SSTab1.Tab = 0
Msbox "Cannot acces Help, Notes or Edit !" & vbCr & "There's no Code selected !", Title, mbOkonly, mbInfo
End If
    
If SSTab1.Tab = 3 Then 'Edit code
Text1(3).SetFocus
End If

End Sub

Private Sub SSTab1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Label4.Tag <> "" Then Exit Sub
If SSTab1.Tab = 0 Then Label4.Caption = "View Code"
If SSTab1.Tab = 1 Then Label4.Caption = "View/Add/Edit Helpfiles"
If SSTab1.Tab = 2 Then Label4.Caption = "View/Add/Edit Notes"
If SSTab1.Tab = 3 Then Label4.Caption = "Edit Code"

End Sub

Private Sub Text1_Change(Index As Integer)
If Index = 1 Then Pic2.BackColor = 0
If Index = 2 Then Pic3.BackColor = 0
If Index = 3 Then Pic4.BackColor = 0
If Text1(0).Text = "" Then
Label10.Caption = ""
Else
Label10.Caption = " Lines of text: " & GetLineCount(Text1(Index)) & " "
End If
End Sub

Private Sub Text1_GotFocus(Index As Integer)
On Error Resume Next
If Index = 0 Then Pic1.SetFocus
End Sub

Private Sub Text1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

   Dim lRetVal As Long
    If Button = vbRightButton Then
        lRetVal = SendMessage(Me.hwnd, WM_RBUTTONDOWN, 0, 0)
  If Index <> 0 Then Exit Sub
  If Text1(0) = "" Then Exit Sub
       Call PopupMenu(mnuCode, 4)
    End If
End Sub

Private Sub Text1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Label4.Tag <> "" Then Exit Sub
If Index = 0 Then Label4.Caption = "Code Window"
If Index = 1 Then Label4.Caption = "Help Window"
If Index = 2 Then Label4.Caption = "Notes Window"
If Index = 3 Then Label4.Caption = "Edit Code Window"
End Sub

Private Sub Toolbar1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Label4.Tag <> "" Then Exit Sub
Label4.Caption = ""
If X > 0 And X < 870 Then Label4.Caption = "Save Database"
If X > 885 And X < 1755 Then Label4.Caption = "Print Code/Helpfile"
If X > 1890 And X < 2760 Then Label4.Caption = "Copy to Clipboard"
If X > 2775 And X < 3645 Then Label4.Caption = "Remove Code/Help/Notes"
If X > 3660 And X < 4530 Then Label4.Caption = "Rename selected Code"
If X > 4545 And X < 5415 Then Label4.Caption = "Move selected Code"
If X > 5550 And X < 6420 Then Label4.Caption = "Add new Category"
If X > 6435 And X < 7305 Then Label4.Caption = "Rename selected Category"
If X > 7320 And X < 8190 Then Label4.Caption = "Remove selected Category"
If X > 8325 And X < 9195 Then Label4.Caption = "Search Code/Helpfiles"
If X > 9330 And X < 10200 Then Label4.Caption = "Help on the CodeLib"

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Pic1.SetFocus
DoEvents
Select Case Button.Key

Case "key3" 'Code to clipboard
mnuCopyCode_Click
'---------------------
Case "key4" 'Delete code
mnuKillCode_Click
'---------------------
Case "key1" 'save database
If Pic5.BackColor <> 0 Then
Msbox "The database has not been changed, so there's no need to save it...", Title, mbOkonly, mbInfo
Exit Sub
End If
Msbox "Save the database ?", Title, mbYesNo, mbQuestion
If MBReturn = 1 Then Exit Sub 'No
SaveLib
Pic5.BackColor = RGB(192, 192, 192)
Label6.Caption = "Database clear"
'---------------------
Case "key2" 'print out code
mnuPrintCod_Click
'---------------------
Case "key5" 'new category
Cat1.Text1.Text = "New Category"
Cat1.Text1.SelLength = Len(Cat1.Text1.Text)
Cat1.Command1.Enabled = False
Cat1.Show 1
'---------------------
Case "key6" 'Rename category
If Label2.Caption = "" Then
Msbox "No category selected...", Title, mbOkonly, mbInfo
Exit Sub
End If
key62:
InBox "Rename the category: " & vbCr & Cat(CatIdx), Cat(CatIdx), Title
If IbReturn = "" Then Exit Sub 'exit
    For xx = 0 To 99
    If LCase(Trim(IbReturn)) = LCase(Cat(xx)) Then
    Msbox "This category already exists !", Title, mbOkonly, mbCritical
    xx = 99
    GoTo key62
    End If
    Next xx
Label2.Caption = Format(CatIdx, "00") & "  " & IbReturn
Cat(CatIdx) = IbReturn
'save categories
On Error GoTo SaveCat2
ff = FreeFile
Open App.Path & "\Data\Cat.ini" For Output As #ff
For xx = 0 To 99
If Cat(xx) = "" Then Exit For
Print #ff, Cat(xx)
Next xx
Close #ff
LoadCat
Exit Sub
SaveCat2:
Close #ff
Msbox "There's an error while" & vbCr & "saving the Category-data..." & vbCr & vbCr & "Error: " & Err & "  " & Err.Description, Title, mbOkonly, mbCritical
Exit Sub
'---------------------
Case "key7" 'rename code
mnuRenameCode_Click
'---------------------
Case "key8" 'Search code
Search.Show 1
'---------------------
Case "key10" 'Move code to another category
mnuMoveCode_Click
'---------------------
Case "key11" 'kill category
If Label2.Caption = "" Then
Msbox "No category selected...", Title, mbOkonly, mbInfo
Exit Sub
End If
For xx = 0 To 999
If CLdata(0, xx) = CatIdx And CLdata(1, xx) <> "" Then
Msbox "This Category '" & Cat(CatIdx) & "' is not empty. You can only remove empty Categories...", Title, mbOkonly, mbCritical
Exit Sub
End If
Next xx
Msbox "Are you sure to remove the selected Category '" & Cat(CatIdx) & "' ?", Title, mbYesNo, mbQuestion
If MBReturn = 1 Then Exit Sub 'No
Cat(CatIdx) = ""
'save categories
On Error GoTo SaveCat3
ff = FreeFile
Open App.Path & "\Data\Cat.ini" For Output As #ff
For xx = 0 To 99
'If Cat(xx) = "" Then Exit For
Print #ff, Cat(xx)
Next xx
Close #ff
LoadCat
Label2.Caption = ""
Label9.Caption = ""
Exit Sub
SaveCat3:
Close #ff
Msbox "There's an error while" & vbCr & "saving the Category-data..." & vbCr & vbCr & "Error: " & Err & "  " & Err.Description, Title, mbOkonly, mbCritical
Exit Sub
'---------------------
Case "key9" 'Helpfiles
HelpScreen.Show 1
End Select
End Sub

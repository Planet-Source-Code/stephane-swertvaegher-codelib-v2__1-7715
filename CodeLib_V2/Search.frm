VERSION 5.00
Begin VB.Form Search 
   AutoRedraw      =   -1  'True
   Caption         =   "CodeLib V2.0"
   ClientHeight    =   5490
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10470
   Icon            =   "Search.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   366
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   698
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C000&
      Caption         =   "Copy as Helpfile"
      Height          =   285
      Left            =   7650
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Back to the CodeLib"
      Top             =   5130
      Width           =   1500
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C000&
      Caption         =   "Select All"
      Height          =   285
      Left            =   4590
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Select all of the text"
      Top             =   5130
      Width           =   1500
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C000&
      Caption         =   "Copy as Code"
      Height          =   285
      Left            =   6120
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Copy selected text"
      Top             =   5130
      Width           =   1500
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   3525
      Left            =   4590
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   7
      ToolTipText     =   "Select text to copy"
      Top             =   1440
      Width           =   5730
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C000&
      Caption         =   "EXIT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   9360
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Back to the CodeLib"
      Top             =   5130
      Width           =   960
   End
   Begin VB.FileListBox File1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H00C00000&
      Height          =   3735
      Left            =   2610
      TabIndex        =   3
      ToolTipText     =   "Double-click to enter the file"
      Top             =   1440
      Width           =   1770
   End
   Begin VB.DirListBox Dir1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H00C00000&
      Height          =   3240
      Left            =   135
      TabIndex        =   2
      Top             =   1980
      Width           =   2265
   End
   Begin VB.DriveListBox Drive1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H00C00000&
      Height          =   315
      Left            =   135
      TabIndex        =   1
      Top             =   1440
      Width           =   2265
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Selected Path: "
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
      TabIndex        =   5
      Top             =   1035
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Search for code"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   510
      Index           =   1
      Left            =   630
      TabIndex        =   4
      Top             =   150
      Width           =   9000
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Search for code"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   510
      Index           =   0
      Left            =   675
      TabIndex        =   0
      Top             =   180
      Width           =   9000
   End
End
Attribute VB_Name = "Search"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click() 'exit
Search.Hide
End Sub

Private Sub Command2_Click() 'copy as code
If Text1.SelText = Empty Then
Msbox "No text selected !", Title, mbOkonly, mbInfo
Exit Sub
End If
SelCode = Text1.SelText
Search2.Show 1
End Sub

Private Sub Command3_Click() 'select all
Text1.SetFocus
Text1.SelStart = 0
Text1.SelLength = Len(Text1.Text)
End Sub

Private Sub Command4_Click() 'copy as helpfile
If Text1.SelText = Empty Then
Msbox "No text selected !", Title, mbOkonly, mbInfo
Exit Sub
End If
SelCode = Text1.SelText
Search3.Show 1
End Sub

Private Sub Dir1_Change()
On Error GoTo Dir12
File1.Path = Dir1.Path
T3D Search, Label3, 5, T3dNone, T3dF1
Label3.Caption = " Selected Path: " & Dir1.Path & " "
T3D Search, Label3, 5, T3dRaiseInset, T3dF1
Text1.Text = ""
SelCode = ""
Exit Sub
Dir12:
Msbox "ERROR ! " & vbCr & "Errornumber: " & Err & vbCr & Err.Description, Title, mbOkonly, mbCritical
End Sub

Private Sub Drive1_Change()
On Error GoTo Drive12
Dir1.Path = Left(Drive1.Drive, 2)
Exit Sub
Drive12:
Msbox "ERROR ! " & vbCr & "Errornumber: " & Err & vbCr & Err.Description, Title, mbOkonly, mbCritical
End Sub

Private Sub File1_DblClick()
Dim Best$
 On Error GoTo FileClick1
If Right(Dir1.Path, 1) = "\" Then
Best = Dir1.Path & File1.List(File1.ListIndex)
Else
Best = Dir1.Path & "\" & File1.List(File1.ListIndex)
End If
ff = FreeFile
Open Best For Input As #ff
  Text1.Text = Input(LOF(1), 1)
Close #ff
Exit Sub
FileClick1:
Close #ff
Msbox Err.Description & vbCr & "File too big to load !", Title, mbOkonly, mbCritical

End Sub

Private Sub Form_Activate()
File1.Pattern = "*.txt;*.bas;*.cls;*.frm;*.ctl;*.doc;*.rtf"
Label3.Caption = " Selected Path: " & Dir1.Path & " "
T3D Search, Label3, 5, T3dRaiseInset
If Text1.Text = "" Then
Command2.Enabled = False
Command3.Enabled = False
Command4.Enabled = False
End If
SelCode = ""
Search2.Combo1.Clear
For xx = 0 To 99
If Cat(xx) <> "" Then
Search2.Combo1.AddItem Cat(xx)
End If
Next xx
Search2.Text1.Text = ""
Search2.Combo1.Text = "Categories"
On Error Resume Next
Text1.SetFocus
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Cancel = True
Search.Hide
End Sub

Private Sub Text1_Change()
If Text1.Text = "" Then
Command2.Enabled = False
Command3.Enabled = False
Command4.Enabled = False
Else
Command2.Enabled = True
Command3.Enabled = True
Command4.Enabled = True
End If
End Sub

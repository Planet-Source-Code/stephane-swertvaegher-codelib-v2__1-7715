VERSION 5.00
Begin VB.Form Cat1 
   AutoRedraw      =   -1  'True
   Caption         =   "CodeLib V2.0"
   ClientHeight    =   3480
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "Cat1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   232
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   312
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C000&
      Caption         =   "Exit"
      Height          =   285
      Left            =   3690
      MaskColor       =   &H00C0C000&
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2970
      Width           =   825
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C000&
      Caption         =   "Accept"
      Height          =   285
      Left            =   2655
      MaskColor       =   &H00C0C000&
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2970
      Width           =   870
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00E0E0E0&
      Height          =   240
      Left            =   225
      TabIndex        =   3
      Top             =   2520
      Width           =   4200
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Height          =   1590
      Left            =   2430
      TabIndex        =   1
      Top             =   720
      Width           =   1995
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Present categories"
      Height          =   195
      Left            =   225
      TabIndex        =   2
      Top             =   720
      Width           =   1995
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Adding a new category"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   330
      Left            =   225
      TabIndex        =   0
      Top             =   180
      Width           =   4200
   End
End
Attribute VB_Name = "Cat1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Text
Private Sub Command1_Click()
For xx = 0 To 99
If Trim(Text1.Text) = Cat(xx) Then
xx = 99
GoTo SaveCat3:
End If
Next xx
For xx = 0 To 99
If Cat(xx) = "" Then
Cat(xx) = Trim(Text1.Text)
Exit For
End If
Next xx
'Save categories
On Error GoTo SaveCat2
ff = FreeFile
Open App.Path & "\Data\Cat.ini" For Output As #ff
For xx = 0 To 99
'If Cat(xx) = "" Then Exit For
Print #ff, Cat(xx)
Next xx
Close #ff
LoadCat
Cat1.Hide
Exit Sub
SaveCat2:
Close #ff
Msbox "There's an error while" & vbCr & "saving the category-data..." & vbCr & vbCr & "Error: " & Err & "  " & Err.Description, Title, mbOkonly, mbCritical
Exit Sub
SaveCat3:
Msbox "This category already exists !", Title, mbOkonly, mbCritical
Text1.SelStart = 0
Text1.SelLength = Len(Text1.Text)
Text1.SetFocus
Cat1.Command1.Enabled = False
End Sub

Private Sub Command2_Click()
Cat1.Hide
End Sub

Private Sub Form_Activate()
List1.Clear
For xx = 0 To 99
If Cat(xx) <> "" Then
List1.AddItem Format(xx, "00") & "  " & Cat(xx)
End If
Next xx
On Error Resume Next
Text1.SetFocus
End Sub

Private Sub Form_Load()
T3D Cat1, Cat1.Label1, 5, T3dRaiseRaise
T3D Cat1, Cat1.Label2, 5, T3dRaiseRaise
T3D Cat1, Cat1.Text1, 5, T3dRaiseRaise
T3D Cat1, Cat1.List1, 5, T3dRaiseRaise

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Cancel = True
Cat1.Hide
End Sub

Private Sub Text1_Change()
If Text1.Text <> "" Then
Command1.Enabled = True
Else
Command1.Enabled = False
End If
End Sub

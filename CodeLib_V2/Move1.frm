VERSION 5.00
Begin VB.Form Move1 
   AutoRedraw      =   -1  'True
   Caption         =   "CodeLib V2.0"
   ClientHeight    =   1755
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4305
   Icon            =   "Move1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   117
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   287
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
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
      Height          =   330
      Left            =   1125
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1395
      Width           =   870
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C000&
      Caption         =   "Accept"
      Height          =   330
      Left            =   135
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1395
      Width           =   870
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H00C00000&
      Height          =   315
      ItemData        =   "Move1.frx":030A
      Left            =   1035
      List            =   "Move1.frx":030C
      TabIndex        =   1
      Text            =   "Combo1"
      Top             =   900
      Width           =   2220
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   510
      Left            =   180
      TabIndex        =   0
      Top             =   135
      Width           =   3930
   End
End
Attribute VB_Name = "Move1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click() 'accept
If Combo1.ListIndex = -1 Then
Msbox "No Category selected !", Title, mbOkonly, mbInfo
Exit Sub
End If
Msbox "Move the selected Code " & vbCr & CLdata(1, CodeLib.List1.ItemData(Idx1)) & vbCr & "in the Category " & Cat(CatIdx) & vbCr & "to the Category " & Combo1.List(Combo1.ListIndex) & " ?" & vbCr, Title, mbYesNo, mbQuestion
If MBReturn = 1 Then Exit Sub
'OK ! Move
CLdata(0, CodeLib.List1.ItemData(Idx1)) = Combo1.ListIndex
CodeLib.Pic5.BackColor = 0
CodeLib.Label6.Caption = "Database dirty"
CodeLib.List1.Clear
SearchItems
Move1.Hide
End Sub

Private Sub Command2_Click() 'exit
Move1.Hide
End Sub

Private Sub Form_Activate()
Combo1.Clear
For xx = 0 To 99
If Cat(xx) <> "" Then
Combo1.AddItem Cat(xx)
End If
Next xx
Combo1.Text = "Pick a Category"
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Cancel = True
Move1.Hide
End Sub

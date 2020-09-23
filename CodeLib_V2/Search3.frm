VERSION 5.00
Begin VB.Form Search3 
   AutoRedraw      =   -1  'True
   Caption         =   "CodeLib V2.0"
   ClientHeight    =   1905
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4620
   Icon            =   "Search3.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   127
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   308
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C000&
      Caption         =   "Accept"
      Height          =   330
      Left            =   270
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1440
      Width           =   870
   End
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
      Left            =   1305
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1440
      Width           =   870
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H00C00000&
      Height          =   315
      Left            =   225
      Sorted          =   -1  'True
      TabIndex        =   1
      Text            =   "Combo1"
      Top             =   720
      Width           =   4155
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Administration of the Helpfile"
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
      Height          =   240
      Left            =   225
      TabIndex        =   0
      Top             =   135
      Width           =   4155
   End
End
Attribute VB_Name = "Search3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click() 'accept
If Combo1.ListIndex = -1 Then
Msbox "No Codefile selected...", Title, mbOkonly, mbInfo
Exit Sub
End If
AppendIdx = Combo1.ItemData(Combo1.ListIndex)
Msbox "Add the selected text to the database as a helpfile to the Code:" & vbCr & vbCr & "Name: " & CLdata(1, AppendIdx) & vbCr & "Category: " & Cat(Val(CLdata(0, AppendIdx))) & vbCr & vbCr & vbCr & "Is this correct ? " & vbCr & vbCr, Title, mbYesNo, mbQuestion
If MBReturn = 1 Then Exit Sub
AddToDB2

End Sub

Private Sub Command2_Click()
Search3.Hide
End Sub

Private Sub Form_Activate()
Combo1.Clear
For xx = 0 To 999
If CLdata(1, xx) <> "" Then
Combo1.AddItem CLdata(1, xx) & " (" & Cat(Val(CLdata(0, xx))) & ")"
Combo1.ItemData(Combo1.NewIndex) = xx
End If
Next xx
Combo1.Text = "Select a Code"
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Cancel = True
Search3.Hide
End Sub

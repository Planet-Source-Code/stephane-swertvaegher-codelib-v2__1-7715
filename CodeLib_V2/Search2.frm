VERSION 5.00
Begin VB.Form Search2 
   AutoRedraw      =   -1  'True
   Caption         =   "CodeLib V2.0"
   ClientHeight    =   2595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3705
   Icon            =   "Search2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   173
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   247
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      ForeColor       =   &H00C00000&
      Height          =   240
      Left            =   1080
      TabIndex        =   6
      Text            =   "Text1"
      Top             =   675
      Width           =   2445
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
      Top             =   2160
      Width           =   870
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C000&
      Caption         =   "Accept"
      Height          =   330
      Left            =   270
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2160
      Width           =   870
   End
   Begin VB.ComboBox Combo1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H00C00000&
      Height          =   315
      Left            =   1125
      TabIndex        =   0
      Text            =   "Combo1"
      Top             =   1125
      Width           =   2265
   End
   Begin VB.Label Label3 
      Caption         =   "Category:"
      Height          =   240
      Left            =   180
      TabIndex        =   5
      Top             =   1170
      Width           =   705
   End
   Begin VB.Label Label2 
      Caption         =   "Name:"
      Height          =   240
      Left            =   180
      TabIndex        =   4
      Top             =   675
      Width           =   705
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Administration of the Code"
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
      TabIndex        =   3
      Top             =   135
      Width           =   3255
   End
End
Attribute VB_Name = "Search2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click() 'copy code
'text empty ?
If Trim(Text1.Text) = "" Then
Msbox "This copy has no name !", Title, mbOkonly, mbInfo
Exit Sub
End If
'name exists ?
For xx = 0 To 999
If LCase(CLdata(1, xx)) = LCase(Trim(Text1.Text)) Then
    ReplaceIdx = xx
    Msbox "The name " & Text1.Text & " in category " & Cat(CLdata(0, xx)) & " already exists !" & vbCr & vbCr & "Would you like to replace the Code ?" & vbCr, Title, mbYesNo, mbInfo
    If MBReturn = 1 Then Exit Sub
    ReplaceCode
    Exit Sub
End If
Next xx
'category exists ?
For xx = 0 To 99
If LCase(Cat(xx)) = LCase(Trim(Combo1.Text)) Then GoTo Command11
Next xx
Msbox "The category " & Combo1.Text & " doesn't exist !", Title, mbOkonly, mbInfo
Exit Sub
Command11:
Msbox "Add the selected text to the database as a Code" & vbCr & vbCr & "Name: " & Text1.Text & vbCr & "Category: " & Combo1.Text & vbCr & vbCr & vbCr & "Is this correct ? " & vbCr & vbCr, Title, mbYesNo, mbQuestion
If MBReturn = 1 Then Exit Sub
AddToDB
End Sub

Private Sub Command2_Click()
Search2.Hide
End Sub

Private Sub Form_Activate()
On Error Resume Next
Text1.SelStart = 0
Text1.SelLength = Len(Text1.Text)
Text1.SetFocus
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Cancel = True
End Sub

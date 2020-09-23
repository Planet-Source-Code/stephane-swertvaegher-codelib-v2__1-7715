VERSION 5.00
Begin VB.Form HelpScreen 
   AutoRedraw      =   -1  'True
   Caption         =   "CodeLib V2.0 - Help"
   ClientHeight    =   5670
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7095
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   24
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "HelpScreen.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   378
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   473
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   225
      Picture         =   "HelpScreen.frx":030A
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   2
      Top             =   270
      Width           =   480
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   4380
      Left            =   135
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Text            =   "HelpScreen.frx":0BD4
      Top             =   1215
      Width           =   6855
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "SendMail"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   5835
      MouseIcon       =   "HelpScreen.frx":0BDA
      MousePointer    =   99  'Custom
      TabIndex        =   3
      ToolTipText     =   "Send e-mail to the coder"
      Top             =   405
      Width           =   825
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   600
      Left            =   1620
      TabIndex        =   1
      Top             =   225
      Width           =   3435
   End
End
Attribute VB_Name = "HelpScreen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Sub Label3_Click()
Dim iret As Long
    Label3.ForeColor = RGB(160, 0, 160)
    iret = ShellExecute(Me.hwnd, vbNullString, "mailto:stephan.swertvaegher@planetinternet.be", vbNullString, "c:\", SW_SHOWNORMAL)
End Sub

Private Sub Rel3D(Obj As Object, Txt$, CX%, CY%, Sh As Long, Fc As Long)
Dim KL%
Obj.ForeColor = Sh
For KL = 1 To 3
Obj.CurrentX = CX + KL
Obj.CurrentY = CY + KL
Obj.Print Txt
Next KL
Obj.ForeColor = Fc
Obj.CurrentX = CX
Obj.CurrentY = CY
Obj.Print Txt
End Sub

Private Sub Form_Load()
Rel3D HelpScreen, "CodeLib V2.0", 125, 15, &HA00000, &HFF8000
ff = FreeFile
    On Error GoTo Load2
    Open App.Path & "\data\Help.txt" For Input As #ff
    Text1.Text = Input(LOF(ff), 1)
Load2:
    Close #ff
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Cancel = True
HelpScreen.Hide
End Sub

Private Sub Text1_GotFocus()
On Error Resume Next
Picture1.SetFocus
End Sub

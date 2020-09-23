Attribute VB_Name = "CLMod1"
Public Idx1%, TabIdx%, CodeCount%
Public CLdata(4, 999), Cat$(99)
Public Temp$, Title$, xx%, yy%, t%, ff%, CLnum%, CatIdx%
Public SelCode$, AppendIdx%, ReplaceIdx%
Public Declare Function SendMessageAsLong Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'Declarations for ExplodeForm
Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type
Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long  'note error in declare
Declare Function Rectangle Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Const WM_RBUTTONDOWN = &H204

Public Sub SizeCombo(frm As Form, cbo As ComboBox)
    Dim cbo_left As Integer
    Dim cbo_top As Integer
    Dim cbo_width As Integer
    Dim cbo_height As Integer
    Dim old_scale_mode As Integer
' Change the Scale Mode on the form to Pixels.
    old_scale_mode = frm.ScaleMode
    frm.ScaleMode = vbPixels
' Save the ComboBox's Left, Top, and Width values.
    cbo_left = cbo.Left
    cbo_top = cbo.Top
    cbo_width = cbo.Width
' Set the new height of the combo box.
    cbo_height = 300
    frm.ScaleMode = old_scale_mode
' Resize the combo box window.
    MoveWindow cbo.hwnd, cbo_left, cbo_top, cbo_width, cbo_height, 1
End Sub


Public Function CryptText(CrTxt$, CrCode)
Dim CrX%
CrCode = CrCode And &HFF& ' max 255
For CrX = 1 To Len(CrTxt)
If Mid(CrTxt, CrX, 1) <> Chr(13) Then
Mid(CrTxt, CrX, 1) = Chr(Asc(Mid(CrTxt, CrX, 1)) Xor CrCode)
End If
Next CrX
CryptText = CrTxt
End Function

Sub ExplodeForm(frm As Form, Steps As Long, Color As Long)
   Dim ThisRect As RECT, RectWidth As Integer, RectHeight As Integer, ScreenDevice As Long, NewBrush As Long, OldBrush As Long, I As Long, X As Integer, Y As Integer, XRect As Integer, YRect As Integer
   If Steps < 20 Then Steps = 20
   'Zooming speed will be different based on machine speed!
   If Color = 0 Then
      Color = frm.BackColor
   End If
   Steps = Steps * 10
   'Get current form window dimensions
   GetWindowRect frm.hwnd, ThisRect
   RectWidth = (ThisRect.Right - ThisRect.Left)
   RectHeight = ThisRect.Bottom - ThisRect.Top
   'Get a device handle for the screen
   ScreenDevice = GetDC(0)
   'Create a brush for drawing to the screen
   'and save the old brush
   NewBrush = CreateSolidBrush(Color)
   OldBrush = SelectObject(ScreenDevice, NewBrush)
   For I = 1 To Steps
      XRect = RectWidth * (I / Steps)
      YRect = RectHeight * (I / Steps)
      X = ThisRect.Left + (RectWidth - XRect) / 2
      Y = ThisRect.Top + (RectHeight - YRect) / 2
      'Incrementally draw rectangle
      Rectangle ScreenDevice, X, Y, X + XRect, Y + YRect
   Next I
   'Return old brush and delete screen device context handle
   'Then destroy brush that drew rectangles
   Call SelectObject(ScreenDevice, OldBrush)
   Call ReleaseDC(0, ScreenDevice)
   DeleteObject (NewBrush)
End Sub

Public Function GetLineCount(C As Control)
  Const EM_GETLINECOUNT = 186
  GetLineCount = SendMessageAsLong(C.hwnd, EM_GETLINECOUNT, 0, 0)
End Function


Public Function Setline(Obj As Object, LineY%, Optional LineStyle As Boolean)
If IsMissing(LineStyle) Then LineStyle = False
If LineStyle = False Then
Obj.Line (0, LineY)-(Obj.ScaleWidth, LineY), RGB(128, 128, 128)
Obj.Line (0, LineY + 1)-(Obj.ScaleWidth, LineY + 1), RGB(240, 240, 240)
Else
Obj.Line (0, LineY)-(Obj.ScaleWidth, LineY), RGB(240, 240, 240)
Obj.Line (0, LineY + 1)-(Obj.ScaleWidth, LineY + 1), RGB(128, 128, 128)
End If
End Function

Public Sub LoadCat()
On Error GoTo LoadCat2
CodeLib.Combo1.Clear
ff = FreeFile
Open App.Path & "\Data\Cat.ini" For Input As #ff
xx = 0
Do While Not EOF(1)

Line Input #ff, Cat(xx)
If Cat(xx) <> Empty Then
CodeLib.Combo1.AddItem Format(xx, "00") & "  " & Cat(xx)
CodeLib.Combo1.ItemData(CodeLib.Combo1.NewIndex) = xx
End If
xx = xx + 1
Loop
Close #ff
Exit Sub
LoadCat2:
Close #ff
Msbox "There's an error while" & vbCr & "loading the category-data..." & vbCr & vbCr & "Error: " & Err & "  " & Err.Description, Title, mbOkonly, mbCritical
End Sub

Public Sub SaveLib()
On Error GoTo SaveLib2
ff = FreeFile
Open App.Path & "\Data\CodeLib.cod" For Output As #ff
For xx = 0 To 999
If CLdata(1, xx) <> "" Then 'has name
    Print #ff, CLdata(0, xx) 'category
    Print #ff, CLdata(1, xx) 'name
    Print #ff, Trim(CLdata(2, xx))
    Print #ff, "÷÷÷÷÷÷" 'code
    If Trim(CLdata(3, xx)) <> "" Then
    Print #ff, Trim(CLdata(3, xx)) 'helpfile
    End If
    Print #ff, "÷÷÷÷÷÷"
    If Trim(CLdata(4, xx)) <> "" Then
    Print #ff, Trim(CLdata(4, xx)) 'notes
    End If
    Print #ff, "÷÷÷÷÷÷"
    
End If
Next xx
Close #ff
Exit Sub
SaveLib2:
Close #ff
Msbox "There's an error while" & vbCr & "saving the Database..." & vbCr & vbCr & "Error: " & Err & "  " & Err.Description, Title, mbOkonly, mbCritical
End Sub

Public Sub LoadLib()
With CodeLib
ff = FreeFile
t = 0
Open App.Path & "\Data\CodeLib.cod" For Input As #ff
'On Error GoTo LoadLib2
Do While Not EOF(1)
Line Input #ff, CLdata(0, t) 'category
Line Input #ff, CLdata(1, t) 'name
    'Load code
    CLdata(2, t) = ""
    Do
    Line Input #1, Temp 'code
    If Temp = "÷÷÷÷÷÷" Then GoTo LoadLib3
    CLdata(2, t) = CLdata(2, t) & Temp & vbCrLf
    Loop
LoadLib3:
    'kill last chr(13) and chr(10)
    If CLdata(2, t) <> "" Then
    CLdata(2, t) = Left(CLdata(2, t), Len(CLdata(2, t)) - 2)
    End If
    'Load Help
    CLdata(3, t) = ""
    Do
    Line Input #1, Temp 'helpfile
    If Temp = "÷÷÷÷÷÷" Then GoTo LoadLib4
    CLdata(3, t) = CLdata(3, t) & Temp & vbCrLf
    Loop
LoadLib4:
    'kill last chr(13) and chr(10)
    If CLdata(3, t) <> "" Then
    CLdata(3, t) = Left(CLdata(3, t), Len(CLdata(3, t)) - 2)
    End If
    'Load Notes
    CLdata(4, t) = ""
    Do
    Line Input #1, Temp 'notes
    If Temp = "÷÷÷÷÷÷" Then GoTo LoadLib5
    CLdata(4, t) = CLdata(4, t) & Temp & vbCrLf
    Loop
LoadLib5:
    'kill last chr(13) and chr(10)
    If CLdata(4, t) <> "" Then
    CLdata(4, t) = Left(CLdata(4, t), Len(CLdata(4, t)) - 2)
    End If
    CodeCount = CodeCount + 1
t = t + 1
Loop
CodeCount = t
Close #ff
Exit Sub
LoadLib2:
Close #ff
Msbox "There's an error while" & vbCr & "loading the database..." & vbCr & vbCr & "Error: " & Err & "  " & Err.Description, Title, mbOkonly, mbCritical
End With
End Sub

Public Sub SearchItems()
For xx = 0 To 999
If CLdata(0, xx) = CatIdx And CLdata(1, xx) <> Empty Then
CodeLib.List1.AddItem CLdata(1, xx)
CodeLib.List1.ItemData(CodeLib.List1.NewIndex) = xx
End If
Next xx
CodeLib.Label9.Caption = CodeLib.List1.ListCount & " items in DataBase"
End Sub

Public Sub RenameCode()
With CodeLib
CLdata(1, .List1.ItemData(Idx1)) = IbReturn
.Label1(0).Caption = .Label2.Caption & " " & CLdata(1, .List1.ItemData(Idx1))
.Label1(1).Caption = .Label2.Caption & " " & CLdata(1, .List1.ItemData(Idx1))
.Label1(2).Caption = .Label2.Caption & " " & CLdata(1, .List1.ItemData(Idx1))
.Pic5.BackColor = 0
.Label6.Caption = "Database dirty"

t = .List1.ListIndex
.List1.Clear
SearchItems
.List1.Selected(t) = True
End With
End Sub
Public Sub KillEntry()
With CodeLib
Screen.MousePointer = 11
For xx = 0 To 4
CLdata(xx, .List1.ItemData(Idx1)) = ""
Next xx
For xx = 0 To 3
.Text1(xx).Text = ""
.Label1(xx).Caption = ""
Next xx

For yy = Idx1 To 998
    For xx = 0 To 4
    CLdata(xx, yy) = CLdata(xx, yy + 1)
    Next xx
Next yy
CLdata(0, 999) = "" 'kill category
CLdata(1, 999) = "" 'kill name
CLdata(2, 999) = "" 'kill code
CLdata(3, 999) = "" 'kill helpfile
CLdata(4, 999) = "" 'kill notes
.Pic5.BackColor = 0
.Label6.Caption = "Database dirty"
.List1.Clear
SearchItems
.Pic7.Visible = False
.Label7.Visible = False
.Pic8.Visible = False
.Label8.Visible = False
DoEvents
.Pic2.BackColor = RGB(192, 192, 192)
.Pic3.BackColor = RGB(192, 192, 192)
.Pic4.BackColor = RGB(192, 192, 192)
CodeCount = CodeCount - 1
.Label11.Caption = "Number of Code-snippets:" & vbCr & CodeCount
Screen.MousePointer = 1
End With
End Sub

Public Sub ColBar(Obj As Object, St%, h%, R%, G%, B%, RE%, GE%, BE%)
Dim H2%, H3%, IvR%, IvG%, IvB%
Obj.AutoRedraw = True
Obj.ScaleMode = 3 'pixel
H3 = Int(h / 2)
IvR = Int(RE - R) / H3
IvG = Int(GE - G) / H3
IvB = Int(BE - B) / H3
Do While h >= H3
Obj.Line (0, St + H2)-(Obj.ScaleWidth, St + H2), RGB(R, G, B)
Obj.Line (0, St + h)-(Obj.ScaleWidth, St + h), RGB(R, G, B)
h = h - 1
H2 = H2 + 1
R = R + IvR
G = G + IvG
B = B + IvB
Loop
End Sub
Public Sub ColBox(Obj As Object, BX%, BY%, EX%, EY%, h%, R%, G%, B%, RE%, GE%, BE%)
Dim H2%, H3%, IvR%, IvG%, IvB%
Obj.AutoRedraw = True
Obj.ScaleMode = 3 'pixel
H3 = Int(h / 2)
IvR = Int(RE - R) / H3
IvG = Int(GE - G) / H3
IvB = Int(BE - B) / H3
Do While h >= H3
Obj.Line (BX + H2, BY + H2)-(EX - H2, EY - H2), RGB(R, G, B), B
Obj.Line (BX + h, BY + h)-(EX - h, EY - h), RGB(R, G, B), B
h = h - 1
H2 = H2 + 1
R = R + IvR
G = G + IvG
B = B + IvB
Loop
End Sub

Public Sub AddToDB() 'add code
    For xx = 0 To 999
    If CLdata(0, xx) = "" And CLdata(1, xx) = "" Then 'no category and name
    AppendIdx = xx
    Exit For
    End If
    Next xx
        CLdata(0, AppendIdx) = Search2.Combo1.ListIndex 'category
        CLdata(1, AppendIdx) = Search2.Text1.Text 'name
        CLdata(2, AppendIdx) = SelCode 'add code

CodeLib.Pic5.BackColor = 0 'database dirty
CodeLib.Label6.Caption = "Database dirty"
If CodeLib.List1.ListCount <> 0 Then
CodeLib.List1.Clear
SearchItems
End If
CodeCount = CodeCount + 1
CodeLib.Label11.Caption = "Number of Code-snippets:" & vbCr & CodeCount
Search2.Hide
End Sub

Public Sub AddToDB2() 'add helpfile
If CLdata(3, AppendIdx) <> "" Then
Msbox "The code " & CLdata(1, AppendIdx) & " has already a helpfile. Would you like to replace it with the new one ?", Title, mbYesNo, mbQuestion
If MBReturn = 1 Then Exit Sub 'do not replace help
End If
        CLdata(3, AppendIdx) = SelCode 'add/replace helpfile
CodeLib.Pic5.BackColor = 0 'database dirty
CodeLib.Label6.Caption = "Database dirty"
Search3.Hide
End Sub

Public Sub ReplaceCode()
CLdata(2, ReplaceIdx) = SelCode ' Replace the code
CodeLib.Pic5.BackColor = 0 'database dirty
CodeLib.Label6.Caption = "Database dirty"
CodeLib.List1.Clear
SearchItems
Search2.Hide
End Sub

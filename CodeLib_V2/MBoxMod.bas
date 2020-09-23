Attribute VB_Name = "MBoxMod"
Public MBReturn%

Public Enum MBoxStyle
 mbOkonly
 mbOKNoWay
 mbOKCancel
 mbYesNo
 mbExitNoWay
 mbSaveNoWay
 mbLoadNoWay
 mbPrintNoWay
 mbEnterLeave
 mbIAgreeLeave
End Enum

Public Enum IconValue
 mbNoIcon
 mbQuestion
 mbInfo
 mbNoEntry
 mbExclamation
 mbSave
 mbOpen
 mbPrint
 mbCritical
 mbTrash
End Enum

 Public Function Msbox(Message As Variant, Optional Title As Variant, Optional Buttons As MBoxStyle, Optional MBoxIcon As IconValue, Optional mbX As Variant, Optional mbY As Variant)
On Error Resume Next
If IsMissing(Title) Then Title = App.Title
' set default
    MBox.Width = 4005 ' 267 pixels
    MBox.Height = 1800 '120 pixels
    MBox.Label2.Width = 200
    MBox.Label2.Height = 13
   MBox.But1(1).Visible = False
   MBox.Label3(1).Caption = ""
    
If Buttons = mbOkonly Then
   MBox.Label3(0).Caption = "OK"
End If
If Buttons = mbOKNoWay Then
   MBox.Label3(0).Caption = "OK"
   MBox.Label3(1).Caption = "No way !"
End If
If Buttons = mbOKCancel Then
   MBox.Label3(0).Caption = "OK"
   MBox.Label3(1).Caption = "Cancel"
End If
If Buttons = mbYesNo Then
   MBox.Label3(0).Caption = "Yes"
   MBox.Label3(1).Caption = "No"
End If
If Buttons = mbExitNoWay Then
   MBox.Label3(0).Caption = "Exit"
   MBox.Label3(1).Caption = "No way !"
End If
If Buttons = mbSaveNoWay Then
   MBox.Label3(0).Caption = "Save"
   MBox.Label3(1).Caption = "No way !"
End If
If Buttons = mbLoadNoWay Then
   MBox.Label3(0).Caption = "Load"
   MBox.Label3(1).Caption = "No way !"
End If
If Buttons = mbPrintNoWay Then
   MBox.Label3(0).Caption = "Print"
   MBox.Label3(1).Caption = "No way !"
End If
If Buttons = mbEnterLeave Then
   MBox.Label3(0).Caption = "Enter"
   MBox.Label3(1).Caption = "Leave"
End If
If Buttons = mbIAgreeLeave Then
   MBox.Label3(0).Caption = "I agree"
   MBox.Label3(1).Caption = "Leave"
End If

If MBox.Label3(1).Caption <> "" Then MBox.But1(1).Visible = True

MBox.Image1.Picture = MBox.ImageList1.ListImages(MBoxIcon).Picture
MBox.Label2.Caption = Message

If MBox.Label2.Width > 200 Then
MBox.Width = (MBox.Label2.Width * 15) + 1005
End If
If MBox.Label2.Height > 46 Then
MBox.Height = (MBox.Label2.Height * 15) + 1110
End If
MBox.Label1.AutoSize = True
MBox.Label1 = Title
MBox.Label1.Left = 12
MBox.Label1.Top = 10
If MBox.Label1.Width > MBox.ScaleWidth - 55 Then
    MBox.Label1.AutoSize = False
    MBox.Label1.Width = MBox.ScaleWidth - 55
End If
MBox.But1(0).Top = MBox.ScaleHeight - 33
MBox.But1(1).Top = MBox.ScaleHeight - 33
    
    MBox.Image1.Left = MBox.ScaleWidth - 40
    MBox.Image1.Top = (MBox.ScaleHeight / 2) - (MBox.Image3.Height / 2)

If MBox.Label3(1).Caption = "" Then
MBox.But1(0).Left = (MBox.ScaleWidth / 2) - (MBox.But1(0).Width / 2)
Else
MBox.But1(0).Left = (MBox.ScaleWidth / 2) - MBox.But1(0).Width - 4
MBox.But1(1).Left = (MBox.ScaleWidth / 2) + 4
End If
MBox.Label3(0).Left = 3
MBox.Label3(1).Left = 3
MBox.Label3(0).Top = 6
MBox.Label3(2).Top = 6
Call PointBar(MBox, 0, 0, 96)
MBox.Line (1, 1)-(MBox.ScaleWidth - 2, MBox.ScaleHeight - 2), RGB(0, 196, 255), B
MBox.Line (1, MBox.ScaleHeight - 2)-(MBox.ScaleWidth - 2, MBox.ScaleHeight - 2), RGB(0, 128, 196)
MBox.Line (MBox.ScaleWidth - 2, 2)-(MBox.ScaleWidth - 2, MBox.ScaleHeight - 1), RGB(0, 128, 196)

If IsMissing(mbX) Then mbX = (Screen.Width / 2) - (MBox.Width / 2)
If IsMissing(mbY) Then mbY = (Screen.Height / 2) - (MBox.Height / 2)
MBox.Left = mbX
MBox.Top = mbY

MBox.Show 1
End Function

Public Sub PointBar(Obj As Object, R%, G%, B%)
Dim Step As Variant, NewStep As Variant, NewR%, NewG%, NewB%, mbT%
Step = 3
NewR = R
NewG = G
NewB = B
mbT = 4
For xx = 0 To 12
Obj.Line (25 - NewStep, mbT + xx)-(Obj.ScaleWidth - 25 + NewStep, mbT + xx), RGB(NewR, NewG, NewB)
Obj.Line (25 - NewStep, mbT + 25 - xx)-(Obj.ScaleWidth - 25 + NewStep, mbT + 25 - xx), RGB(NewR, NewG, NewB)
NewStep = NewStep + Step
NewR = NewR + 10
If NewR > 255 Then NewR = 255
NewG = NewG + 10
If NewG > 255 Then NewG = 255
NewB = NewB + 10
If NewB > 255 Then NewB = 255
Step = Step - 0.25
Next xx
End Sub


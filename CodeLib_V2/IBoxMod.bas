Attribute VB_Name = "IBoxMod"
Public IbReturn, aa As String

Public Function InBox(Message As Variant, Optional IPrev As Variant, Optional Title As Variant, Optional IbX As Variant, Optional IbY As Variant) ' As Integer 'IboxResult
'fixed length & width InputBox
'Message = max. 5 lines of text
On Error Resume Next
If IsMissing(Title) Then Title = App.Title
IBox.Label1 = Title
IBox.Label1.Left = 12
IBox.Label1.Top = 10

IBox.Label2.Caption = Message
If IsMissing(IPrev) Then IPrev = ""
IBox.Text1.Text = IPrev
IBox.Text1.SelLength = Len(IBox.Text1.Text)
'setborder
Call PointBar(IBox, 0, 0, 96)
IBox.Line (1, 1)-(IBox.ScaleWidth - 2, IBox.ScaleHeight - 2), RGB(0, 196, 255), B
IBox.Line (1, IBox.ScaleHeight - 2)-(IBox.ScaleWidth - 2, IBox.ScaleHeight - 2), RGB(0, 128, 196)
IBox.Line (IBox.ScaleWidth - 2, 2)-(IBox.ScaleWidth - 2, IBox.ScaleHeight - 1), RGB(0, 128, 196)

If IsMissing(IbX) Then IbX = (Screen.Width / 2) - (IBox.Width / 2)
If IsMissing(IbY) Then IbY = (Screen.Height / 2) - (IBox.Height / 2)
IBox.Left = IbX
IBox.Top = IbY


IBox.Show 1
End Function


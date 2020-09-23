Attribute VB_Name = "T3DMod"
Public Enum T3dFill
T3dF0
T3dF1
End Enum

Public Enum Borderstyle
T3dRaiseRaise
T3dRaiseInset
T3dInsetRaise
T3dInsetInset
T3dNone
End Enum

Public Function T3D(Obj0 As Object, Obj As Object, Bev%, Optional Style3D As Borderstyle, Optional T3dFilled As T3dFill)
Dim R1%, G1%, B1%, R2%, G2%, B2%, R3%, G3%, B3%, R4%, G4%, B4%
Dim T3Dxx%
On Error Resume Next

Obj.Borderstyle = 0 'no border

If IsMissing(Style3D) Then Style3D = 0

If Style3D > 4 Then Style3D = 3

If Style3D = 0 Then 'RaiseRaise
R1 = 240: R2 = 128: R3 = 240: R4 = 128
End If
If Style3D = 1 Then 'RaiseInset
R1 = 240: R2 = 128: R4 = 240: R3 = 128
End If
If Style3D = 2 Then 'InsetRaise
R2 = 240: R1 = 128: R3 = 240: R4 = 128
End If
If Style3D = 3 Then 'InsetInset
R2 = 240: R1 = 128: R4 = 240: R3 = 128
End If
If Style3D = 4 Then 'No Border
R1 = 192: R2 = 192: R3 = 192: R4 = 192
End If

G1 = R1: B1 = R1
G2 = R2: B2 = R2
G3 = R3: B3 = R3
G4 = R4: B4 = R4
Bev = Bev + 1
T3Dxx = Bev
'Outer
If IsMissing(T3dFilled) Or T3dFilled = 0 Then
    Obj0.Line (Obj.Left - Bev, Obj.Top - Bev)-(Obj.Left - Bev, Obj.Top + Obj.Height + Bev), RGB(R1, G1, B1)
    Obj0.Line (Obj.Left - Bev, Obj.Top - Bev)-(Obj.Left + Obj.Width + Bev, Obj.Top - Bev), RGB(R1, G1, B1)
    Obj0.Line (Obj.Left + Obj.Width + Bev, Obj.Top - Bev)-(Obj.Left + Obj.Width + Bev, Obj.Top + Obj.Height + Bev), RGB(R2, G2, B2)
    Obj0.Line (Obj.Left - Bev, Obj.Top + Obj.Height + Bev)-(Obj.Left + Obj.Width + Bev + 1, Obj.Top + Obj.Height + Bev), RGB(R2, G2, B2)
Else
For Bev = T3Dxx To 1 Step -1
    Obj0.Line (Obj.Left - Bev, Obj.Top - Bev)-(Obj.Left - Bev, Obj.Top + Obj.Height + Bev), RGB(R1, G1, B1)
    Obj0.Line (Obj.Left - Bev, Obj.Top - Bev)-(Obj.Left + Obj.Width + Bev, Obj.Top - Bev), RGB(R1, G1, B1)
    Obj0.Line (Obj.Left + Obj.Width + Bev, Obj.Top - Bev)-(Obj.Left + Obj.Width + Bev, Obj.Top + Obj.Height + Bev), RGB(R2, G2, B2)
    Obj0.Line (Obj.Left - Bev, Obj.Top + Obj.Height + Bev)-(Obj.Left + Obj.Width + Bev + 1, Obj.Top + Obj.Height + Bev), RGB(R2, G2, B2)
Next Bev
End If
'Inner
    Obj0.Line (Obj.Left - 1, Obj.Top - 1)-(Obj.Left - 1, Obj.Top + Obj.Height + 1), RGB(R3, G3, B3)
    Obj0.Line (Obj.Left - 1, Obj.Top - 1)-(Obj.Left + Obj.Width + 1, Obj.Top - 1), RGB(R3, G3, B3)
    Obj0.Line (Obj.Left + Obj.Width + 1, Obj.Top - 1)-(Obj.Left + Obj.Width + 1, Obj.Top + Obj.Height + 1), RGB(R4, G4, B4)
    Obj0.Line (Obj.Left - 1, Obj.Top + Obj.Height + 1)-(Obj.Left + Obj.Width + 2, Obj.Top + Obj.Height + 1), RGB(R4, G4, B4)
End Function


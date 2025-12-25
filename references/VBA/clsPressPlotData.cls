VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsPressPlotData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False


Option Explicit
Option Base 1
Option Compare Binary

Private lngOutZ As Long, lngOutD As Long
Private dblOutR(2) As Double
Private dblOutCX(2) As Double, dblOutCY(2) As Double
Private dblOutSH(2) As Double
Private dblOutCZ(2, 2) As Double, dblOutCH(2, 2) As Double


Private Sub Class_Initialize()

Dim j As Long

lngOutZ = 0: lngOutD = 0
For j = 1 To 2
  dblOutR(j) = 10000000000#
  dblOutCX(j) = 0: dblOutCY(j) = 0
  dblOutSH(j) = 0
  dblOutCZ(1, j) = 0: dblOutCH(1, j) = 0
  dblOutCZ(2, j) = 0: dblOutCH(2, j) = 0
Next

End Sub


Public Property Get lngZ() As Long

lngZ = lngOutZ

End Property


Public Property Get lngD() As Long

lngD = lngOutD

End Property


Public Property Get dblRadius(ByVal lngRn As Long) As Double

dblRadius = dblOutR(lngRn)

End Property


Public Property Get dblCenterX(ByVal lngRn As Long) As Double

dblCenterX = dblOutCX(lngRn)

End Property


Public Property Get dblCenterY(ByVal lngRn As Long) As Double

dblCenterY = dblOutCY(lngRn)

End Property


Public Property Get dblSurfaceH(ByVal lngRn As Long) As Double

dblSurfaceH = dblOutSH(lngRn)

End Property


Public Property Get dblCornerZ(ByVal lngEn As Long, ByVal lngRn As Long) As Double

dblCornerZ = dblOutCZ(lngEn, lngRn)

End Property


Public Property Get dblCornerH(ByVal lngEn As Long, ByVal lngRn As Long) As Double

dblCornerH = dblOutCH(lngEn, lngRn)

End Property


Public Sub Calc()

Dim objPressParam As clsPressParam

Dim i As Long, j As Long
Dim stdCX As Double, stdCY As Double, zVal(7) As Double
Dim d(2) As Double, c(2) As Double, h(2) As Double

If Not IsNumeric(Cells(8, 3).Value) Or _
   Cells(8, 3).Value = 0 Then
  dblOutR(1) = 10000000000#
Else
  dblOutR(1) = Cells(8, 3).Value
End If
If Not IsNumeric(Cells(9, 3).Value) Or _
   Cells(9, 3).Value = 0 Then
  dblOutR(2) = 10000000000#
Else
  dblOutR(2) = Cells(9, 3).Value
End If

zVal(1) = Abs(Cells(13, 3).Value / dblOutR(1) - _
              Cells(14, 3).Value / dblOutR(2)) / 4
zVal(2) = 0.04
zVal(3) = 0.06
zVal(4) = 0.1
zVal(5) = 0.15
zVal(6) = 0.2
zVal(7) = 10000000000#
For i = 2 To UBound(zVal)
  If zVal(1) < zVal(i) Then
    lngOutZ = i - 1: Exit For
  End If
Next

If Cells(10, 3).Value / Cells(8, 6).Value < 0.4 Then
  lngOutD = 7
Else
  lngOutD = 8
End If

Set objPressParam = New clsPressParam
objPressParam.Import strConfig:="PressParam"
With objPressParam
  For j = 1 To 2
    For i = 1 To .lngMaxCount("DP¡–@")
      If Cells(8, 6).Value < .dbldDP(i, 9) Then
        dblOutCH(2, j) = Cells(8, 6).Value / 2 + .dbldDP(i, lngOutZ) / 2
        d(j) = .dbldDP(i, lngOutD) / 2
        Exit For
      End If
    Next
    
    dblOutR(j) = dblOutR(j) + d(j) * (j * (-2) + 3)
    For i = 1 To .lngMaxCount("RŒ…")
      If Abs(dblOutR(j)) / dblOutCH(2, j) < .dblrDP(i, 9) Then
        If .dblrDP(i, 1) = 0 Then
          dblOutR(j) = 10000000000#
        Else
          dblOutR(j) = Fix(dblOutR(j) / .dblrDP(i, 1)) * .dblrDP(i, 1)
          If dblOutR(j) * (j * (-2) + 3) > 0 Then
            dblOutR(j) = dblOutR(j) + .dblrDP(i, 1) * (j * (-2) + 3)
          End If
        End If
        Exit For
      End If
    Next
    
    For i = 1 To .lngMaxCount("C–ÊŽæ")
      If dblOutCH(2, j) * 2 < .dblcDP(i, 9) Then
        c(j) = .dblcDP(i, 1): Exit For
      End If
    Next
  Next
End With
Set objPressParam = Nothing

With rpdDoc.DrawInfo
  If .Sheet = zwSheetA3 Then
    stdCX = 111
    stdCY = 168
  Else
    stdCX = 105
    stdCY = 169
  End If
  dblOutCX(1) = stdCX / .Numerator - Cells(10, 3).Value / 2 - d(1)
  dblOutCX(2) = stdCX / .Numerator + Cells(10, 3).Value / 2 + d(2)
  dblOutCY(1) = stdCY / .Numerator
  dblOutCY(2) = stdCY / .Numerator
End With

If dblOutR(1) > 0 And _
   dblOutR(1) / (dblOutCH(2, 1) * 2) < 3 Then
  c(1) = 0
End If
If dblOutR(2) < 0 And _
   dblOutR(2) / (dblOutCH(2, 2) * 2) > -3 Then
  c(2) = 0
End If
dblOutCH(1, 1) = dblOutCH(2, 1) - c(1)
dblOutCH(1, 2) = dblOutCH(2, 2) - c(2)
h(1) = Cells(11, 6).Value - d(1) + 0.5
h(2) = Cells(12, 6).Value - d(2) + 0.5
If h(1) > 0 And _
   Abs(dblOutR(1)) < 1000000000# And _
   Cells(8, 6).Value > Cells(9, 6).Value Then
  dblOutSH(1) = Sqr(2 * Abs(dblOutR(1)) * h(1) - h(1) ^ 2)
Else
  dblOutSH(1) = dblOutCH(1, 1)
End If
If h(2) > 0 And _
   Abs(dblOutR(2)) < 1000000000# And _
   Cells(8, 6).Value > Cells(10, 6).Value Then
  dblOutSH(2) = Sqr(2 * Abs(dblOutR(2)) * h(2) - h(2) ^ 2)
Else
  dblOutSH(2) = dblOutCH(1, 2)
End If
dblOutCZ(1, 1) = Sag(Radius:=dblOutR(1), Diameter:=dblOutSH(1) * 2)
dblOutCZ(1, 2) = Sag(Radius:=dblOutR(2), Diameter:=dblOutSH(2) * 2)
dblOutCZ(2, 1) = dblOutCZ(1, 1) + c(1)
dblOutCZ(2, 2) = dblOutCZ(1, 2) - c(2)

End Sub


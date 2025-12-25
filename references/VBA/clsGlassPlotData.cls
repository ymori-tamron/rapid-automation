VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsGlassPlotData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False


Option Explicit
Option Base 1
Option Compare Binary

Private blnOutA(2) As Boolean
Private lngOutMC As Long
Private dblOutR(2) As Double
Private dblOutCX(2) As Double, dblOutCY(2) As Double
Private dblOutSH(2) As Double, dblOutDH(2) As Double
Private dblOutCZ(2, 2) As Double, dblOutCH(2, 2) As Double
Private dblOutZ(0 To 50, 2) As Double


Private Sub Class_Initialize()

Dim i As Long, j As Long

lngOutMC = 0
For j = 1 To 2
  blnOutA(j) = False
  dblOutR(j) = 10000000000#
  dblOutCX(j) = 0: dblOutCY(j) = 0
  dblOutSH(j) = 0: dblOutDH(j) = 0
  dblOutCZ(1, j) = 0: dblOutCH(1, j) = 0
  dblOutCZ(2, j) = 0: dblOutCH(2, j) = 0
  For i = 0 To UBound(dblOutZ, 1)
    dblOutZ(i, j) = 0
  Next
Next

End Sub


Public Property Get lngMaxCount() As Long

lngMaxCount = lngOutMC

End Property


Public Property Get blnAsp(ByVal lngRn As Long) As Boolean

blnAsp = blnOutA(lngRn)

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


Public Property Get dblDeltaH(ByVal lngRn As Long) As Double

dblDeltaH = dblOutDH(lngRn)

End Property


Public Property Get dblCornerZ(ByVal lngEn As Long, ByVal lngRn As Long) As Double

dblCornerZ = dblOutCZ(lngEn, lngRn)

End Property


Public Property Get dblCornerH(ByVal lngEn As Long, ByVal lngRn As Long) As Double

dblCornerH = dblOutCH(lngEn, lngRn)

End Property


Public Property Get dblSagZ(ByVal lngCount As Long, ByVal lngRn As Long) As Double

dblSagZ = dblOutZ(lngCount, lngRn)

End Property


Public Sub Calc()

Dim i As Long
Dim stdCX As Double, stdCY As Double
Dim divNum(2) As Double

If Not IsEmpty(Cells(19, 3).Value) Then
  blnOutA(1) = True
Else
  blnOutA(1) = False
End If
If Not IsEmpty(Cells(26, 3).Value) Then
  blnOutA(2) = True
Else
  blnOutA(2) = False
End If

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

With rpdDoc.DrawInfo
  If .Sheet = zwSheetA3 Then
    stdCX = 111
    stdCY = 168
  Else
    stdCX = 105
    stdCY = 169
  End If
  dblOutCX(1) = stdCX / .Numerator - Cells(10, 3).Value / 2
  dblOutCX(2) = stdCX / .Numerator + Cells(10, 3).Value / 2
  dblOutCY(1) = stdCY / .Numerator
  dblOutCY(2) = stdCY / .Numerator
End With

If Not CBool(fCheck(3)) Then
  lngOutMC = UBound(dblOutZ, 1)
Else
  With rpdDoc.DrawInfo
    If Cells(8, 6).Value > Cells(9, 6).Value Then
      divNum(1) = (Cells(9, 6).Value / 2) / (2 / .Numerator)
    Else
      divNum(1) = (Cells(9, 6).Value / 2 - Cells(13, 6).Value) / (2 / .Numerator)
    End If
    If Cells(8, 6).Value > Cells(10, 6).Value Then
      divNum(2) = (Cells(10, 6).Value / 2) / (2 / .Numerator)
    Else
      divNum(2) = (Cells(10, 6).Value / 2 - Cells(14, 6).Value) / (2 / .Numerator)
    End If
  End With
  
  If divNum(1) <= divNum(2) Then
    lngOutMC = Fix(divNum(1))
  Else
    lngOutMC = Fix(divNum(2))
  End If
  
  If lngOutMC < 10 Then
    lngOutMC = 10
  ElseIf lngOutMC > UBound(dblOutZ, 1) Then
    lngOutMC = UBound(dblOutZ, 1)
  End If
End If

If Cells(8, 6).Value > Cells(9, 6).Value Then
  dblOutSH(1) = Cells(9, 6).Value / 2
  dblOutDH(1) = (Cells(9, 6).Value / 2) / lngOutMC
Else
  dblOutSH(1) = 0
  dblOutDH(1) = (Cells(9, 6).Value / 2 - Cells(13, 6).Value) / lngOutMC
End If
If Cells(8, 6).Value > Cells(10, 6).Value Then
  dblOutSH(2) = Cells(10, 6).Value / 2
  dblOutDH(2) = (Cells(10, 6).Value / 2) / lngOutMC
Else
  dblOutSH(2) = 0
  dblOutDH(2) = (Cells(10, 6).Value / 2 - Cells(14, 6).Value) / lngOutMC
End If
For i = 0 To lngOutMC
  dblOutZ(i, 1) = Sag(Radius:=Cells(8, 3).Value, Diameter:=(dblOutDH(1) * i) * 2, _
                      Conic:=Cells(19, 3).Value, _
                      A4:=Cells(20, 3).Value, A6:=Cells(21, 3).Value, _
                      A8:=Cells(22, 3).Value, A10:=Cells(23, 3).Value, _
                      A12:=Cells(24, 3).Value, A14:=Cells(25, 3).Value)
  dblOutZ(i, 2) = Sag(Radius:=Cells(9, 3).Value, Diameter:=(dblOutDH(2) * i) * 2, _
                      Conic:=Cells(26, 3).Value, _
                      A4:=Cells(27, 3).Value, A6:=Cells(28, 3).Value, _
                      A8:=Cells(29, 3).Value, A10:=Cells(30, 3).Value, _
                      A12:=Cells(31, 3).Value, A14:=Cells(32, 3).Value)
Next

dblOutCZ(1, 1) = dblOutZ(lngOutMC, 1)
dblOutCZ(1, 2) = dblOutZ(lngOutMC, 2)
dblOutCH(1, 1) = Cells(8, 6).Value / 2 - Cells(13, 6).Value
dblOutCH(1, 2) = Cells(8, 6).Value / 2 - Cells(14, 6).Value
dblOutCZ(2, 1) = dblOutZ(lngOutMC, 1) + Cells(13, 6).Value
dblOutCZ(2, 2) = dblOutZ(lngOutMC, 2) - Cells(14, 6).Value
dblOutCH(2, 1) = Cells(8, 6).Value / 2
dblOutCH(2, 2) = Cells(8, 6).Value / 2

End Sub


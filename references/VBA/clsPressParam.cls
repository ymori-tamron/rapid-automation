VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsPressParam"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False


Option Explicit
Option Base 1
Option Compare Binary

Private lngOutMC(6) As Long
Private dblOutdDP(99, 9) As Double, dblOutdHP(99, 9) As Double
Private dblOuttDP(99, 9) As Double, dblOuttHP(99, 9) As Double
Private dblOutcDP(99, 9) As Double, dblOutrDP(99, 9) As Double

Private strItem(99) As String


Private Sub Class_Initialize()

Dim i As Long, j As Long

For i = 1 To UBound(lngOutMC)
  lngOutMC(i) = 0
Next

For i = 1 To UBound(strItem)
  strItem(i) = ""
  For j = 1 To 9
    dblOutdDP(i, j) = 0: dblOutdHP(i, j) = 0
    dblOuttDP(i, j) = 0: dblOuttHP(i, j) = 0
    dblOutcDP(i, j) = 0: dblOutrDP(i, j) = 0
  Next
Next

End Sub


Public Property Get lngMaxCount(ByVal strType As String) As Long

Select Case strType
  Case "DP寸法": lngMaxCount = lngOutMC(1)
  Case "HP寸法": lngMaxCount = lngOutMC(2)
  Case "DP公差": lngMaxCount = lngOutMC(3)
  Case "HP公差": lngMaxCount = lngOutMC(4)
  Case "C面取": lngMaxCount = lngOutMC(5)
  Case "R桁": lngMaxCount = lngOutMC(6)
End Select

End Property


Public Property Get dbldDP(ByVal lngIndex1 As Long, ByVal lngIndex2 As Long) As Double

dbldDP = dblOutdDP(lngIndex1, lngIndex2)

End Property


Public Property Get dbldHP(ByVal lngIndex1 As Long, ByVal lngIndex2 As Long) As Double

dbldHP = dblOutdHP(lngIndex1, lngIndex2)

End Property


Public Property Get dbltDP(ByVal lngIndex1 As Long, ByVal lngIndex2 As Long) As Double

dbltDP = dblOuttDP(lngIndex1, lngIndex2)

End Property


Public Property Get dbltHP(ByVal lngIndex1 As Long, ByVal lngIndex2 As Long) As Double

dbltHP = dblOuttHP(lngIndex1, lngIndex2)

End Property


Public Property Get dblcDP(ByVal lngIndex1 As Long, ByVal lngIndex2 As Long) As Double

dblcDP = dblOutcDP(lngIndex1, lngIndex2)

End Property


Public Property Get dblrDP(ByVal lngIndex1 As Long, ByVal lngIndex2 As Long) As Double

dblrDP = dblOutrDP(lngIndex1, lngIndex2)

End Property


Public Sub Import(ByVal strConfig As String)

Dim f As Integer
Dim i As Long, j As Long, lngIndex As Long
Dim tmp(9) As Double
Dim cfgFileName As String

cfgFileName = ThisWorkbook.Path & "\Config\" & strConfig & ".cfg"
f = FreeFile(0)
Open cfgFileName For Input As #f
For i = 1 To UBound(strItem)
  If Not EOF(f) Then
    Input #f, lngIndex, strItem(i), _
              tmp(1), tmp(2), tmp(3), tmp(4), tmp(5), _
              tmp(6), tmp(7), tmp(8), tmp(9)
    Select Case strItem(i)
      Case "DP寸法"
        lngOutMC(1) = lngOutMC(1) + 1
        For j = 1 To 9
          dblOutdDP(lngOutMC(1), j) = tmp(j)
        Next
      Case "HP寸法"
        lngOutMC(2) = lngOutMC(2) + 1
        For j = 1 To 9
          dblOutdHP(lngOutMC(2), j) = tmp(j)
        Next
      Case "DP公差"
        lngOutMC(3) = lngOutMC(3) + 1
        For j = 1 To 9
          dblOuttDP(lngOutMC(3), j) = tmp(j)
        Next
      Case "HP公差"
        lngOutMC(4) = lngOutMC(4) + 1
        For j = 1 To 9
          dblOuttHP(lngOutMC(4), j) = tmp(j)
        Next
      Case "C面取"
        lngOutMC(5) = lngOutMC(5) + 1
        For j = 1 To 9
          dblOutcDP(lngOutMC(5), j) = tmp(j)
        Next
      Case "R桁"
        lngOutMC(6) = lngOutMC(6) + 1
        For j = 1 To 9
          dblOutrDP(lngOutMC(6), j) = tmp(j)
        Next
    End Select
  Else
    Exit For
  End If
Next
Close #f

End Sub


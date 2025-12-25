VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsRapidParam"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False


Option Explicit
Option Base 1
Option Compare Binary

Private lngOutMC As Long
Private lngOutVal(99, 2) As Long
Private dblOutXY(99, 2) As Double, dblOutHW(99, 2) As Double
Private strOutItem(99) As String


Private Sub Class_Initialize()

Dim i As Long, j As Long

lngOutMC = 0
For i = 1 To UBound(strOutItem)
  strOutItem(i) = ""
  For j = 1 To 2
    lngOutVal(i, j) = 0
    dblOutXY(i, j) = 0
    dblOutHW(i, j) = 0
  Next
Next

End Sub


Public Property Get lngMaxCount() As Long

lngMaxCount = lngOutMC

End Property


Public Property Get strItem(ByVal lngIndex As Long) As String

strItem = strOutItem(lngIndex)

End Property


Public Property Get dblXY(ByVal lngIndex1 As Long, ByVal lngIndex2 As Long) As Double

dblXY = dblOutXY(lngIndex1, lngIndex2)

End Property


Public Property Get dblHW(ByVal lngIndex1 As Long, ByVal lngIndex2 As Long) As Double

dblHW = dblOutHW(lngIndex1, lngIndex2)

End Property


Public Property Get lngVal(ByVal lngIndex1 As Long, ByVal lngIndex2 As Long) As Long

lngVal = lngOutVal(lngIndex1, lngIndex2)

End Property


Public Sub Import(ByVal strConfig As String)

Dim f As Integer
Dim i As Long, lngIndex As Long
Dim cfgFileName As String

cfgFileName = ThisWorkbook.Path & "\Config\" & strConfig & ".cfg"
f = FreeFile(0)
Open cfgFileName For Input As #f
For i = 1 To UBound(strOutItem)
  If Not EOF(f) Then
    Input #f, lngIndex, strOutItem(i), dblOutXY(i, 1), dblOutXY(i, 2), _
              dblOutHW(i, 1), dblOutHW(i, 2), lngOutVal(i, 1), lngOutVal(i, 2)
  Else
    lngOutMC = i - 1: Exit For
  End If
Next
Close #f

End Sub


Public Sub SetUp(ByVal strParam As String)

Select Case strParam
  Case "DimeParam"
    With rpdDoc.DimensionParam
      .FaceName = strOutItem(1)
      .Height = dblOutHW(1, 1)
      .Width = dblOutHW(1, 2)
      .DecimalDigit = lngOutVal(1, 1)
      .ZeroSuppress = CBool(lngOutVal(1, 2))
      .StartMarkType = zwMarkArrow
      .EndMarkType = zwMarkArrow
      .ArrowLength = dblOutXY(1, 1)
      .CircleRadius = dblOutXY(1, 2)
      .LegDistance = dblOutXY(1, 1) / 2
      .TopDistance = dblOutXY(1, 1) / 2
      .StartDistance = dblOutXY(1, 1) * 2
      .EndDistance = dblOutXY(1, 1) * 2
    End With
  Case "PickParam"
    With rpdDoc.PickupParam
      .FaceName = strOutItem(1)
      .Height = dblOutHW(1, 1)
      .Width = dblOutHW(1, 2)
      .MarkType = zwMarkArrow
      .ArrowLength = dblOutXY(1, 1)
      .CircleRadius = dblOutXY(1, 2)
    End With
End Select

End Sub


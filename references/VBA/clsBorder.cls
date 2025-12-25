VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsBorder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False


Option Explicit
Option Base 1
Option Compare Binary


Public Sub SetUp(ByRef rngCells As Range)

With rngCells
  .Borders(xlDiagonalDown).LineStyle = xlNone
  .Borders(xlDiagonalUp).LineStyle = xlNone
  With .Borders(xlEdgeLeft)
    .LineStyle = xlContinuous: .Weight = xlHairline
  End With
  With .Borders(xlEdgeTop)
    .LineStyle = xlContinuous: .Weight = xlHairline
  End With
  With .Borders(xlEdgeBottom)
    .LineStyle = xlContinuous: .Weight = xlHairline
  End With
  With .Borders(xlEdgeRight)
    .LineStyle = xlContinuous: .Weight = xlHairline
  End With
  
  If .Rows.Count > 1 Then
    With .Borders(xlInsideHorizontal)
      .LineStyle = xlContinuous: .Weight = xlHairline
    End With
  End If
  If .Columns.Count > 1 Then
    With .Borders(xlInsideVertical)
      .LineStyle = xlContinuous: .Weight = xlHairline
    End With
  End If
End With

End Sub


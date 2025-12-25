VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsFormula"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False


Option Explicit
Option Base 1
Option Compare Binary


Public Sub SetUp(ByRef rngCell As Range, ByVal strFormula As String)

With rngCell
  .Formula = strFormula
  .Interior.ColorIndex = 34
End With

End Sub


VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsSagTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False


Option Explicit
Option Base 1
Option Compare Binary

Private lngCount As Long
Private dblParam(99, 9) As Double
Private strTitle(99) As String


Private Sub Class_Initialize()

Dim i As Long, j As Long

lngCount = 0
For i = 1 To UBound(strTitle)
  strTitle(i) = ""
  For j = 1 To 9
    dblParam(i, j) = 0
  Next
Next

End Sub


Public Sub PickUp()

If Not IsEmpty(Cells(19, 3).Value) Then
  lngCount = lngCount + 1
  If Not IsNumeric(Cells(8, 3).Value) Or _
     Cells(8, 3).Value = 0 Then
    dblParam(lngCount, 1) = 10000000000#
  Else
    dblParam(lngCount, 1) = Cells(8, 3).Value
  End If
  dblParam(lngCount, 2) = Fix((Cells(9, 6).Value / 2) / 0.1) * 0.1
  dblParam(lngCount, 3) = Cells(19, 3).Value
  dblParam(lngCount, 4) = Cells(20, 3).Value
  dblParam(lngCount, 5) = Cells(21, 3).Value
  dblParam(lngCount, 6) = Cells(22, 3).Value
  dblParam(lngCount, 7) = Cells(23, 3).Value
  dblParam(lngCount, 8) = Cells(24, 3).Value
  dblParam(lngCount, 9) = Cells(25, 3).Value
  strTitle(lngCount) = "G" & Mid(ActiveSheet.Name, 2, 2) & "-R1"
End If

If Not IsEmpty(Cells(26, 3).Value) Then
  lngCount = lngCount + 1
  If Not IsNumeric(Cells(9, 3).Value) Or _
     Cells(9, 3).Value = 0 Then
    dblParam(lngCount, 1) = 10000000000#
  Else
    dblParam(lngCount, 1) = Cells(9, 3).Value
  End If
  dblParam(lngCount, 2) = Fix((Cells(10, 6).Value / 2) / 0.1) * 0.1
  dblParam(lngCount, 3) = Cells(26, 3).Value
  dblParam(lngCount, 4) = Cells(27, 3).Value
  dblParam(lngCount, 5) = Cells(28, 3).Value
  dblParam(lngCount, 6) = Cells(29, 3).Value
  dblParam(lngCount, 7) = Cells(30, 3).Value
  dblParam(lngCount, 8) = Cells(31, 3).Value
  dblParam(lngCount, 9) = Cells(32, 3).Value
  strTitle(lngCount) = "G" & Mid(ActiveSheet.Name, 2, 2) & "-R2"
End If

End Sub


Public Sub Export()

Dim i As Long, j As Long, cc As Long
Dim dh As Double
Dim fml(3) As String

If lngCount > 0 Then
  Workbooks.Open _
    Filename:=ThisWorkbook.Path & "\Template\サグデータ.xltx"
  DoEvents
  
  Cells(2, 2).Value = mdlNum & " - 非球面形状"
  Cells(3, 3).Value = "(1/" & CStr(Fix((lngCount - 1) / 4) + 1) & ")"
  Cells(2, 12).Value = Date
  
  If Fix((lngCount - 1) / 4) > 0 Then
    For i = 1 To Fix((lngCount - 1) / 4)
      cc = i * 13
      ActiveSheet.VPageBreaks.Add Before:=Cells(1, cc + 14)
      
      Range(Columns(1), Columns(13)).Copy
      Range(Columns(cc + 1), Columns(cc + 13)).PasteSpecial _
        Paste:=xlPasteColumnWidths
      Application.CutCopyMode = False
      
      Range(Cells(2, 2), Cells(3, 12)).Copy _
        Destination:=Range(Cells(2, cc + 2), Cells(3, cc + 12))
      Cells(3, cc + 3).Value = "(" & CStr(i + 1) & "/" & _
                               CStr(Fix((lngCount - 1) / 4) + 1) & ")"
    Next
  End If
  
  For i = 1 To lngCount
    cc = (i - 1) * 3 + 2 + Fix((i - 1) / 4)
    If i > 1 Then
      Range(Cells(5, 2), Cells(72, 3)).Copy
      Range(Cells(5, cc), Cells(72, cc + 1)).PasteSpecial _
        Paste:=xlPasteFormats
      Application.CutCopyMode = False
      
      Range(Cells(6, 2), Cells(14, 2)).Copy _
        Destination:=Range(Cells(6, cc), Cells(14, cc))
      Range(Cells(16, 2), Cells(16, 3)).Copy _
        Destination:=Range(Cells(16, cc), Cells(16, cc + 1))
    End If
    
    Cells(5, cc).Value = strTitle(i)
    Cells(6, cc + 1).Value = dblParam(i, 1)
    Cells(7, cc + 1).Value = dblParam(i, 3) + 1
    Cells(8, cc + 1).Value = 0
    Cells(9, cc + 1).Value = dblParam(i, 4)
    Cells(10, cc + 1).Value = dblParam(i, 5)
    Cells(11, cc + 1).Value = dblParam(i, 6)
    Cells(12, cc + 1).Value = dblParam(i, 7)
    Cells(13, cc + 1).Value = dblParam(i, 8)
    Cells(14, cc + 1).Value = dblParam(i, 9)
    
    fml(1) = _
      Cells(6, cc + 1).Address(RowAbsolute:=False, ColumnAbsolute:=False)
    fml(3) = "," & _
      Cells(7, cc + 1).Address(RowAbsolute:=False, ColumnAbsolute:=False) & "-1"
    fml(3) = fml(3) & "," & _
      Cells(9, cc + 1).Address(RowAbsolute:=False, ColumnAbsolute:=False)
    fml(3) = fml(3) & "," & _
      Cells(10, cc + 1).Address(RowAbsolute:=False, ColumnAbsolute:=False)
    fml(3) = fml(3) & "," & _
      Cells(11, cc + 1).Address(RowAbsolute:=False, ColumnAbsolute:=False)
    fml(3) = fml(3) & "," & _
      Cells(12, cc + 1).Address(RowAbsolute:=False, ColumnAbsolute:=False)
    fml(3) = fml(3) & "," & _
      Cells(13, cc + 1).Address(RowAbsolute:=False, ColumnAbsolute:=False)
    fml(3) = fml(3) & "," & _
      Cells(14, cc + 1).Address(RowAbsolute:=False, ColumnAbsolute:=False)
    
    dh = Fix(((dblParam(i, 2) - 0.1) / 52) / 0.1) * 0.1 + 0.1
    For j = 0 To 54
      If j * dh < dblParam(i, 2) + dh * 2 Then
        Cells(j + 17, cc).Value = j * dh
        fml(2) = "," & _
          Cells(j + 17, cc).Address(RowAbsolute:=False, ColumnAbsolute:=False)
        Cells(j + 17, cc + 1).Formula = _
          "=Sag(" & fml(1) & fml(2) & "*2" & fml(3) & ")"
      Else
        Exit For
      End If
    Next
  Next
  
  ActiveSheet.DisplayPageBreaks = True
  Cells(1, 1).Activate
End If

End Sub


VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsToleranceData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False


Option Explicit
Option Base 1
Option Compare Binary


Public Sub AddG()

Dim objFormula As clsFormula

Dim cc As Long, maxCol As Long
Dim nt As String, rt As String
Dim at As String, sf As String
Dim crt As String, tmp As String

Set objFormula = New clsFormula

Worksheets(DefaultSheet3).Activate
maxCol = Cells(6, Columns.Count).End(xlToLeft).Column
If maxCol = 1 Then
  cc = 3
  Cells(2, 3).Value = mdlNum
  Cells(3, 3).Value = Date
Else
  cc = maxCol + 3
  Columns(cc - 1).ColumnWidth = 0.5
  Range(Columns(cc), Columns(cc + 1)).ColumnWidth = 8
  Columns(cc + 2).ColumnWidth = 2
  
  Range(Cells(6, 3), Cells(26, 4)).Copy
  With Range(Cells(6, cc), Cells(26, cc + 1))
    .PasteSpecial Paste:=xlPasteFormats
    .Interior.ColorIndex = xlNone
  End With
  Application.CutCopyMode = False
End If

With Worksheets(Worksheets.Count - 1)
  Cells(6, cc).Value = .Name
  If Not IsEmpty(.Cells(13, 3).Value) Then
    Cells(7, cc).Value = 3
    Cells(9, cc).Value = 1
  End If
  If Not IsEmpty(.Cells(14, 3).Value) Then
    Cells(8, cc).Value = 3
    Cells(10, cc).Value = 1
  End If
  
  If Not IsEmpty(.Cells(19, 3).Value) Then
    nt = Cells(7, cc).Address(RowAbsolute:=False, ColumnAbsolute:=False)
    rt = Cells(7, cc + 1).Address(RowAbsolute:=False, ColumnAbsolute:=False)
    at = Cells(9, cc).Address(RowAbsolute:=False, ColumnAbsolute:=False)
    If .Cells(8, 3).Value < 0 Then
      sf = Chr(34) & "CC" & Chr(34)
    Else
      sf = Chr(34) & "CX" & Chr(34)
    End If
    objFormula.SetUp rngCell:=Cells(7, cc + 1), strFormula:= _
      "=NewtonToDeltaRadius(" & nt & "," & sf & ",'" & .Name & "'!C8," & _
                                                 "'" & .Name & "'!C13)"
    objFormula.SetUp rngCell:=Cells(9, cc + 1), _
                     strFormula:="=" & rt & "*" & at & "/" & nt
  End If
  If Not IsEmpty(.Cells(26, 3).Value) Then
    nt = Cells(8, cc).Address(RowAbsolute:=False, ColumnAbsolute:=False)
    rt = Cells(8, cc + 1).Address(RowAbsolute:=False, ColumnAbsolute:=False)
    at = Cells(10, cc).Address(RowAbsolute:=False, ColumnAbsolute:=False)
    If .Cells(9, 3).Value > 0 Then
      sf = Chr(34) & "CC" & Chr(34)
    Else
      sf = Chr(34) & "CX" & Chr(34)
    End If
    objFormula.SetUp rngCell:=Cells(8, cc + 1), strFormula:= _
      "=NewtonToDeltaRadius(" & nt & "," & sf & ",'" & .Name & "'!C9," & _
                                                 "'" & .Name & "'!C14)"
    objFormula.SetUp rngCell:=Cells(10, cc + 1), _
                     strFormula:="=" & rt & "*" & at & "/" & nt
  End If
  
  If InStr(.Name, "ASP") > 0 Then
    Cells(11, cc).Value = 0.02
    Cells(11, cc + 1).Value = -0.02
  Else
    Cells(11, cc).Value = 0.03
    Cells(11, cc + 1).Value = -0.03
    If .Cells(10, 9).Value = "-" Or _
       .Cells(11, 9).Value = "-" Then
      If .Cells(14, 9).Value > 0 Then
        Cells(11, cc).Value = 0.02
        Cells(11, cc + 1).Value = -0.04
      End If
    End If
  End If
  Cells(12, cc).Value = 0.0005
  tmp = Cells(12, cc).Address(RowAbsolute:=False, ColumnAbsolute:=False)
  objFormula.SetUp _
    rngCell:=Cells(12, cc + 1), strFormula:="=" & tmp & "*100000"
  Cells(13, cc).Value = 0.008
  tmp = Cells(13, cc).Address(RowAbsolute:=False, ColumnAbsolute:=False)
  objFormula.SetUp _
    rngCell:=Cells(13, cc + 1), strFormula:="=" & tmp & "*'" & .Name & "'!I9"
  
  Select Case Left(.Name, 1)
    Case "G"
      Cells(14, cc).Value = 0.005
      Cells(14, cc + 1).Value = -0.005
      If .Cells(10, 9).Value = "-" Or _
         .Cells(11, 9).Value = "-" Then
        If .Cells(14, 9).Value > 0 Then
          Cells(14, cc).Value = 0.01
          Cells(14, cc + 1).Value = -0.01
        End If
      End If
      Cells(15, cc).Value = "()"
      Cells(16, cc).Value = "()"
      If .Cells(8, 6).Value > .Cells(9, 6).Value Then
        Cells(17, cc).Value = 0.01
        Cells(17, cc + 1).Value = -0.01
      End If
      If .Cells(8, 6).Value > .Cells(10, 6).Value Then
        Cells(18, cc).Value = 0.01
        Cells(18, cc + 1).Value = -0.01
      End If
      Cells(19, cc).Value = ""
      Cells(20, cc).Value = ""
      crt = Cells(25, cc).Address(RowAbsolute:=False, ColumnAbsolute:=False)
      If InStr(.Name, "ASP") > 0 Then
        Cells(21, cc).Value = 1.5
        Cells(22, cc).Value = 1.5
        tmp = Cells(22, cc).Address(RowAbsolute:=False, ColumnAbsolute:=False)
        objFormula.SetUp rngCell:=Cells(22, cc + 1), strFormula:= _
          "=ABS(IF(" & crt & "=1,'" & .Name & "'!C8," & _
                                "'" & .Name & "'!C9))*TAN(RADIANS(" & tmp & "/60))"
        If Not IsEmpty(.Cells(19, 3).Value) Then
          Cells(25, cc).Value = 1
          If Not IsEmpty(Cells(17, cc).Value) Then
            Cells(17, cc).Value = 0.005
            Cells(17, cc + 1).Value = -0.005
          End If
        ElseIf Not IsEmpty(.Cells(26, 3).Value) Then
          Cells(25, cc).Value = 2
          If Not IsEmpty(Cells(18, cc).Value) Then
            Cells(18, cc).Value = 0.005
            Cells(18, cc + 1).Value = -0.005
          End If
        End If
        Cells(26, cc).Value = 5
        Cells(26, cc + 1).Value = 1
      Else
        Cells(21, cc).Value = 1
        If .Cells(8, 6).Value > .Cells(9, 6).Value Or _
           .Cells(8, 6).Value > .Cells(10, 6).Value Then
          Cells(23, cc).Value = 1
        End If
        If .Cells(8, 6).Value > .Cells(9, 6).Value And _
           .Cells(8, 6).Value > .Cells(10, 6).Value Then
          Cells(24, cc).Value = 1
        End If
        tmp = Cells(23, cc).Address(RowAbsolute:=False, ColumnAbsolute:=False)
        objFormula.SetUp rngCell:=Cells(23, cc + 1), strFormula:= _
          "=IF('" & .Name & "'!F8>IF(" & crt & "=1,'" & .Name & "'!F9," & _
                                                  "'" & .Name & "'!F10)," & _
              "'" & .Name & "'!F8*TAN(RADIANS(" & tmp & "/60)),"""")"
        tmp = Cells(24, cc).Address(RowAbsolute:=False, ColumnAbsolute:=False)
        objFormula.SetUp rngCell:=Cells(24, cc + 1), strFormula:= _
          "=IF(AND('" & .Name & "'!F8>'" & .Name & "'!F9," & _
                  "'" & .Name & "'!F8>'" & .Name & "'!F10)," & _
                  "'" & .Name & "'!F8*TAN(RADIANS(" & tmp & "/60)),"""")"
        If .Cells(8, 6).Value > .Cells(9, 6).Value Then
          Cells(25, cc).Value = 1
        ElseIf .Cells(8, 6).Value > .Cells(10, 6).Value Then
          Cells(25, cc).Value = 2
        End If
      End If
      
    Case "H"
      If Not IsEmpty(.Cells(13, 3).Value) Then
        Cells(15, cc).Value = 0.6
        Cells(15, cc + 1).Value = -0.3
      End If
      If Not IsEmpty(.Cells(14, 3).Value) Then
        Cells(16, cc).Value = 0.6
        Cells(16, cc + 1).Value = -0.3
      End If
  End Select
End With
Cells(1, 1).Activate

Set objFormula = Nothing

End Sub


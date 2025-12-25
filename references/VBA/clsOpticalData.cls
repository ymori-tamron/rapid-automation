VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsOpticalData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False


Option Explicit
Option Base 1
Option Compare Binary


Public Sub Add(ByVal opFileName As String)

Dim objBorder As clsBorder

Dim k As Variant
Dim f As Integer
Dim i As Long, j As Long, sc As Long, ic As Long
Dim apc As Long, ddtc As Long, maxCol As Long
Dim fil As String, tmp As String, txt As String
Dim TextLine As String, strAsp(99, 9) As String

Set objBorder = New clsBorder

Worksheets(DefaultSheet1).Activate
If InStrRev(opFileName, "\") > 0 Then
  fil = Mid(opFileName, InStrRev(opFileName, "\") + 1)
Else
  fil = "#NAME?"
End If

f = FreeFile(0)
Open opFileName For Input As #f
maxCol = 12
ddtc = 0
tmp = ""
Do While Not EOF(f)
  Line Input #f, TextLine
  If Right(TextLine, 1) = "&" Then
    tmp = tmp & Trim(Mid(TextLine, 1, Len(TextLine) - 1)) & " "
  Else
    TextLine = tmp & Trim(TextLine): tmp = ""
    
    If Left(TextLine, 5) = "TITLE" Then
      If InStr(8, TextLine, " ") > 0 Then
        mdlNum = Trim(Mid(TextLine, 8, InStr(8, TextLine, " ") - 8))
      Else
        mdlNum = Trim(Mid(TextLine, 8, Len(TextLine) - 8))
      End If
    End If
    
    If Left(TextLine, 2) = "S " Or _
       Left(TextLine, 2) = "SO" Or _
       Left(TextLine, 2) = "SI" Then
      lngSur = lngSur + 1
      With Cells(lngSur, 4)
        .Value = TextLine
        .TextToColumns DataType:=xlDelimited, ConsecutiveDelimiter:=True, Space:=True
        .Value = "S" & CStr(lngSur - 1)
      End With
      
      If Cells(lngSur, 5).Value = 0 Then Cells(lngSur, 5).Value = "Inf"
      If Cells(lngSur, 6).Value > 1000000000# Then Cells(lngSur, 6).Value = "Inf"
      If Not IsEmpty(Cells(lngSur, 7).Value) Then
        txt = UCase(Cells(lngSur, 7).Value)
        If Left(txt, 1) = "'" Then
          Cells(lngSur, 7).Value = Mid(txt, 2, Len(txt) - 2)
          Cells(lngSur, 8).Value = "OTHER"
        Else
          If InStr(txt, "_") > 0 Then
            Cells(lngSur, 7).Value = Mid(txt, 1, InStr(txt, "_") - 1)
            Cells(lngSur, 8).Value = Mid(txt, InStr(txt, "_") + 1)
          End If
        End If
        
        If Cells(lngSur, 7).Value = "AMT121B" Or _
           Cells(lngSur, 7).Value = "AMT122B" Or _
           Cells(lngSur, 7).Value = "MP252" Or _
           Cells(lngSur, 7).Value = "MP281" Then
          Cells(lngSur, 3).Value = "HYB"
        End If
      End If
      
      With Range(Cells(lngSur, 5), Cells(lngSur, 9))
        .NumberFormat = "0.000"
        .HorizontalAlignment = xlCenter
      End With
      
      objBorder.SetUp rngCells:=Range(Cells(lngSur, 2), Cells(lngSur, 10))
    End If
    If Left(TextLine, 3) = "STO" Then
      If Cells(lngSur, 3).Value <> "HYB" Then
        Cells(lngSur, 3).Value = "STO"
      End If
    End If
    If Left(TextLine, 3) = "CIR" Then
      If Left(TextLine, 10) = "CIR L'DEF'" Then
        Cells(lngSur, 9).Value = Trim(Mid(TextLine, 11))
      ElseIf Left(TextLine, 5) Like "CIR #" Then
        Cells(lngSur, 9).Value = Trim(Mid(TextLine, 4))
        Cells(lngSur, 10).Value = "☆"
      End If
    End If
    
    If Left(TextLine, 3) = "ASP" Then
      lngAsp = lngAsp + 1
      strAsp(lngAsp, 1) = Cells(lngSur, 4).Value
      strAsp(lngAsp, 2) = Cells(lngSur, 5).Value
      Cells(lngSur, 4).Value = Cells(lngSur, 4).Value & "(A" & CStr(lngAsp) & ")"
      Cells(lngSur, 5).NumberFormat = "0.0000"
    End If
    If Left(TextLine, 2) = "K " Then
      If InStr(TextLine, ";") > 0 Then
        strAsp(lngAsp, 3) = Trim(Mid(TextLine, 2, InStr(TextLine, ";") - 2))
      Else
        strAsp(lngAsp, 3) = Trim(Mid(TextLine, 2))
      End If
    End If
    If Left(TextLine, 2) = "A " Then
      TextLine = TextLine & ";": sc = 1
      For j = 4 To 7
        If InStr(sc + 3, TextLine, ";") > 0 Then
          strAsp(lngAsp, j) = Trim(Mid(TextLine, sc + 3, InStr(sc + 3, TextLine, ";") - (sc + 3)))
          sc = InStr(sc + 3, TextLine, ";")
        Else
          Exit For
        End If
      Next
    End If
    If Left(TextLine, 2) = "D " Then
      strAsp(lngAsp, 7) = Trim(Mid(TextLine, 2))
    End If
    If Left(TextLine, 2) = "E " Then
      TextLine = TextLine & ";": sc = 1
      For j = 8 To 9
        If InStr(sc + 3, TextLine, ";") > 0 Then
          strAsp(lngAsp, j) = Trim(Mid(TextLine, sc + 3, InStr(sc + 3, TextLine, ";") - (sc + 3)))
          sc = InStr(sc + 3, TextLine, ";")
        Else
          Exit For
        End If
      Next
    End If
    
    If Left(TextLine, 9) = "ZOO   THI" Then
      lngZoo = lngZoo + 1
      With Cells(lngSur + 1 + lngZoo, 4)
        .Value = Mid(TextLine, 12)
        .TextToColumns DataType:=xlDelimited, ConsecutiveDelimiter:=True, Space:=True
        
        With .Offset(0, 1).Resize(1, .End(xlToRight).Column - 4)
          For Each k In .Cells
            If k.Value > 1000000000# Then k.Value = "Inf"
          Next
          
          .NumberFormat = "0.000"
          .HorizontalAlignment = xlCenter
        End With
        
        objBorder.SetUp rngCells:=.Offset(0, -1).Resize(1, .End(xlToRight).Column - 2)
        If .End(xlToRight).Column > maxCol Then maxCol = .End(xlToRight).Column
      End With
    End If
    
    If Left(TextLine, 5) = "! DDT" Then
      ddtc = ddtc + 1
      DDT(ddtc) = Mid(TextLine, 7) & " "
    End If
  End If
Loop
Close #f

If lngSur > 3 Then
  apc = 0
  For i = 2 To lngSur - 2
    If Not IsEmpty(Cells(i, 7).Value) Then
      If Cells(i, 3).Value <> "HYB" Then
        If Cells(i, 5).Value <> "Inf" Or _
           Cells(i + 1, 5).Value <> "Inf" Then
          lngGla = lngGla + 1
          If lngGla < 10 Then
            Cells(i, 3).Value = "G0" & CStr(lngGla)
          Else
            Cells(i, 3).Value = "G" & CStr(lngGla)
          End If
          
          If Left(Cells(i - 1, 3).Value, 1) = "G" Then
            lngBal = lngBal + 1
            If lngBal < 10 Then
              Cells(i - 1, 2).Value = "B0" & CStr(lngBal)
            Else
              Cells(i - 1, 2).Value = "B" & CStr(lngBal)
            End If
          End If
        End If
      End If
    Else
      If IsEmpty(Cells(i - 1, 7).Value) Then
        If Cells(i, 3).Value <> "STO" And _
           Cells(i, 10).Value = "☆" Then
          apc = apc + 1
          Cells(i, 3).Value = "AP" & CStr(apc)
        End If
      End If
    End If
  Next
End If

If lngAsp > 0 Then
  For i = 1 To lngAsp
    ic = lngSur + lngZoo + 2 + i
    Cells(ic, 3).Value = "A" & CStr(i)
    For j = 1 To 9
      With Cells(ic, j + 3)
        If strAsp(i, j) <> "" Then
          .Value = strAsp(i, j)
        Else
          .Value = "0.0"
        End If
        
        If j > 1 Then
          Select Case j
            Case 2: .NumberFormat = "0.0000"
            Case 3: .NumberFormat = "+0.0000;-0.0000"
            Case Is >= 4: .NumberFormat = "+0.00000E+00;-0.00000E+00"
          End Select
          .HorizontalAlignment = xlCenter
        End If
      End With
    Next
    
    objBorder.SetUp rngCells:=Range(Cells(ic, 3), Cells(ic, 12))
  Next
End If

If lngZoo > 0 Then
  ic = lngSur + 1
  Cells(ic + 1, 4).CurrentRegion.Sort Key1:=Cells(ic + 1, 4), Header:=xlNo
  
  For i = 1 To lngZoo
    Cells(ic + i, 3).Value = "D" & CStr(i)
    With Cells(ic + i, 4)
      Cells(.Value + 1, 6).Value = "(D" & CStr(i) & ")"
      .Value = "S" & .Value
    End With
  Next
End If

Columns(1).ColumnWidth = 2
Range(Columns(2), Columns(3)).ColumnWidth = 4.5
Columns(4).ColumnWidth = 8.5
Range(Columns(5), Columns(maxCol)).ColumnWidth = 14
Columns(maxCol + 1).ColumnWidth = 2

Range(Rows(1), Rows(3)).Insert Shift:=xlShiftDown
Cells(2, 4).Value = "機種名:" & mdlNum
Cells(2, 6).Value = "ファイル名:" & fil
Cells(2, 10).Value = Date

If lngSur > 0 Then
  ic = 5
  Range(Cells(ic - 1, 1), Cells(ic, 1)).EntireRow.Insert _
    Shift:=xlShiftDown, CopyOrigin:=xlFormatFromRightOrBelow
  Cells(ic - 1, 2).Value = "■基本データ"
  Cells(ic, 2).Value = "レンズ"
  With Range(Cells(ic, 2), Cells(ic, 3))
    .Merge
    .HorizontalAlignment = xlCenter
  End With
  With Cells(ic, 4)
    .Value = "面番号"
    .HorizontalAlignment = xlCenter
  End With
  Cells(ic, 5).Value = "曲率半径"
  Cells(ic, 6).Value = "面間隔"
  Cells(ic, 7).Value = "硝種"
  Cells(ic, 8).Value = "メーカー名"
  Cells(ic, 9).Value = "有効半径"
  With Cells(ic, 10)
    .Value = "光線カット"
    .HorizontalAlignment = xlCenter
  End With
  
  objBorder.SetUp rngCells:=Range(Cells(ic, 2), Cells(ic, 10))
End If

If lngAsp > 0 Then
  ic = lngSur + lngZoo + 9
  Range(Cells(ic - 1, 1), Cells(ic, 1)).EntireRow.Insert _
    Shift:=xlShiftDown, CopyOrigin:=xlFormatFromRightOrBelow
  Cells(ic - 1, 3).Value = "■非球面データ(非球面式はCODEV形式)"
  With Cells(ic, 4)
    .Value = "面番号"
    .HorizontalAlignment = xlCenter
  End With
  Cells(ic, 5).Value = "曲率半径"
  Cells(ic, 6).Value = "コーニック定数"
  Cells(ic, 7).Value = "4次の係数"
  Cells(ic, 8).Value = "6次の係数"
  Cells(ic, 9).Value = "8次の係数"
  Cells(ic, 10).Value = "10次の係数"
  Cells(ic, 11).Value = "12次の係数"
  Cells(ic, 12).Value = "14次の係数"
  
  objBorder.SetUp rngCells:=Range(Cells(ic, 3), Cells(ic, 12))
End If

If lngZoo > 0 Then
  ic = lngSur + 8
  Range(Cells(ic - 1, 1), Cells(ic, 1)).EntireRow.Insert _
    Shift:=xlShiftDown, CopyOrigin:=xlFormatFromRightOrBelow
  Cells(ic - 1, 3).Value = "■ズームデータ"
  With Cells(ic, 4)
    .Value = "面番号"
    .HorizontalAlignment = xlCenter
    For j = 5 To .CurrentRegion.Columns.Count + 2
      Cells(ic, j).Value = "POS" & CStr(j - 4)
    Next
    
    objBorder.SetUp rngCells:=Range(Cells(ic, 3), _
                                    Cells(ic, .CurrentRegion.Columns.Count + 2))
  End With
End If

Set objBorder = Nothing

End Sub


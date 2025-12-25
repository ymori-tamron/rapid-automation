VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsGlassData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False


Option Explicit
Option Base 1
Option Compare Binary

Private intOut As Integer


Private Sub Class_Initialize()

intOut = 0

End Sub


Public Property Get intHyb() As Integer

intHyb = intOut

End Property


Public Sub AddA()

Dim objFormula As clsFormula
Dim objBorder As clsBorder
Dim Unused As Range

Dim chk(2) As Boolean
Dim i As Long, j As Long, l As Long, m As Long, n As Long
Dim cc As Long, pos As Long
Dim sZoo As Long, eZoo As Long, dZoo As Long
Dim rc(3) As Long, maxRow(2) As Long
Dim zRow(99) As Long, sRow(99) As Long, eRow(99) As Long
Dim maxLen As Double
Dim tmp As String, txt(4) As String

Set objFormula = New clsFormula
Set objBorder = New clsBorder

Columns(5).ColumnWidth = 18
Range(Columns(6), Columns(8)).ColumnWidth = 14
Columns(9).ColumnWidth = 20

Cells(3, 5).Value = "品名(上段)"
Cells(4, 5).Value = "品名(下段)"

Cells(7, 2).Value = "■焦点距離"
Cells(8, 2).Value = ""
Cells(18, 2).Value = "■間隔公差"
Cells(19, 2).Value = ""

With Cells(8, 3)
  .NumberFormat = """f= ""+0.0#;""f= ""-0.0#;""f= Inf"";""f= ""@"
  .HorizontalAlignment = xlRight
End With
With Cells(19, 3)
  .NumberFormat = "@"
  .HorizontalAlignment = xlRight
End With

Set Unused = Union(Range(Cells(9, 2), Cells(14, 3)), _
                   Range(Cells(20, 2), Cells(32, 3)), _
                   Range(Cells(7, 5), Cells(16, 6)), _
                   Range(Cells(30, 5), Cells(36, 6)), _
                   Range(Cells(7, 8), Cells(16, 9)))
With Unused
  .Clear
  .Font.Name = "ＭＳ 明朝"
  .Font.Size = 10
End With
Set Unused = Nothing

With Range(Cells(18, 2), Cells(19, 3))
  .Cut Destination:=Range(Cells(10, 2), Cells(11, 3))
  .Font.Name = "ＭＳ 明朝"
  .Font.Size = 10
End With
With Range(Cells(20, 5), Cells(28, 9))
  .Cut Destination:=Range(Cells(11, 5), Cells(19, 9))
  .Font.Name = "ＭＳ 明朝"
  .Font.Size = 10
End With

Cells(7, 5).Value = "■作図設定"
Cells(8, 5).Value = "ズームポジション"

Cells(20, 5).Value = "注記9"
Cells(21, 5).Value = "注記10"

Cells(23, 5).Value = "■光学仕様書"
Cells(24, 5).Value = "E.F.L.(INF)"
Cells(25, 5).Value = "O.V.L.(INF)/in glass"
Cells(26, 5).Value = "O.V.L.(INF)/in air  "
Cells(27, 5).Value = "F.NO.(INF)"
Cells(28, 5).Value = "絞り径"
Cells(29, 5).Value = "射出瞳位置(INF)"
Cells(30, 5).Value = "最短撮影距離(MOD)"
Cells(31, 5).Value = "最大撮影倍率(MOD)"
Cells(32, 5).Value = "基準テスト位置"
Cells(33, 5).Value = "L(INF)"
Cells(34, 5).Value = "B.F.L.(INF)/in glass"
Cells(35, 5).Value = "B.F.L.(INF)/in air  "
Cells(36, 5).Value = "繰り出し量/基準位置"
Cells(37, 5).Value = "繰り出し量/MOD     "

Cells(8, 5).ShrinkToFit = True
objBorder.SetUp rngCells:=Range(Cells(8, 5), Cells(8, 6))

With Range(Cells(20, 6), Cells(21, 9))
  .ShrinkToFit = True
  .Merge Across:=True
End With
objBorder.SetUp rngCells:=Range(Cells(20, 5), Cells(21, 9))

Cells(23, 8).HorizontalAlignment = xlRight
Range(Cells(24, 5), Cells(37, 5)).ShrinkToFit = True
Range(Cells(24, 6), Cells(37, 8)).HorizontalAlignment = xlRight
Range(Cells(24, 6), Cells(24, 8)).NumberFormat = _
  "0.00"" mm"";-0.00"" mm"";0.00"" mm"";@"" mm"""
Range(Cells(25, 6), Cells(26, 8)).NumberFormat = _
  "0.000"" mm"";-0.000"" mm"";0.000"" mm"";@"" mm"""
Range(Cells(27, 6), Cells(27, 8)).NumberFormat = _
  """1:""0.000""  "";""1:""-0.000""  "";""1:""0.000""  "";""1:""@""  """
Range(Cells(28, 6), Cells(28, 8)).NumberFormat = _
  """φ""0.000"" mm"";""φ""-0.000"" mm"";""φ""0.000"" mm"";""φ""@"" mm"""
Range(Cells(29, 6), Cells(29, 8)).NumberFormat = _
  "+0.00"" mm"";-0.00"" mm"";-0.00"" mm"";""-""@"" mm"""
Range(Cells(30, 6), Cells(30, 8)).NumberFormat = _
  "+0.0"" mm"";-0.0"" mm"";-0.0"" mm"";""-""@"" mm"""
Range(Cells(31, 6), Cells(31, 8)).NumberFormat = _
  """+1/""0.00""  "";""-1/""0.00""  "";""NA  "";""-1/""@""  """
Range(Cells(32, 6), Cells(32, 8)).NumberFormat = _
  "+0.00"" m "";-0.00"" m "";-0.00"" m "";""-""@"" m """
Range(Cells(33, 6), Cells(37, 8)).NumberFormat = _
  "0.000"" mm"";-0.000"" mm"";0.000"" mm"";@"" mm"""
objBorder.SetUp rngCells:=Range(Cells(24, 5), Cells(37, 8))

Cells(2, 3).Value = mdlNum
Cells(4, 3).Value = Date
Cells(3, 6).Value = "(機種の特徴を記入)"

If ActiveSheet.Name <> Alpha(14) Then
  Cells(3, 3).Value = ActiveSheet.Name & "," & Alpha(14)
  Cells(4, 6).Value = "光学配置図,光学仕様書"
Else
  Cells(3, 3).Value = ActiveSheet.Name
  Cells(4, 6).Value = "配置図･仕様書(1/2),配置図･仕様書(2/2)"
End If

If lngZoo > 0 Then
  With Worksheets(DefaultSheet1)
    tmp = ""
    For j = 5 To .Cells(lngSur + 8, 4).End(xlToRight).Column
      If j > 5 Then tmp = tmp & ","
      tmp = tmp & .Cells(lngSur + 8, j).Value
    Next
  End With
  With Cells(8, 6)
    With .Validation
      .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
           Formula1:=tmp
      .InCellDropdown = True
      .IgnoreBlank = False
      .ShowInput = False
      .ShowError = True
      .ErrorTitle = ProjectTitle
      .ErrorMessage = "リスト以外の項目は入力出来ません"
      .IMEMode = xlIMEModeNoControl
    End With
    .Value = "POS1"
  End With
Else
  With Cells(8, 6)
    .Borders(xlDiagonalUp).LineStyle = xlContinuous
    .Borders(xlDiagonalUp).Weight = xlHairline
  End With
End If

Cells(12, 6).Value = "☆は光線カット位置を表す"
If DDTValue("TIT", 1) <> "NaN" Then
  With Worksheets(DefaultSheet1)
    tmp = ""
    For i = DDTValue("FOC", 1) + 6 To DDTValue("FOC", 2) + 6
      If .Cells(i, 3).Value = "STO" Then
        If tmp = "" Then
          tmp = "絞り面"
        Else
          tmp = Left(tmp, 3) & "-" & "絞り面"
        End If
      ElseIf Left(.Cells(i, 3).Value, 1) = "G" Then
        If tmp = "" Then
          tmp = .Cells(i, 3).Value
        Else
          tmp = Left(tmp, 3) & "-" & .Cells(i, 3).Value
        End If
      End If
    Next
    If DDTValue("FCS", 1) > 0 Then
      tmp = tmp & "を対物側に移動させることで行う"
    Else
      tmp = tmp & "を接眼側に移動させることで行う"
    End If
    Cells(13, 6).Value = "フォーカスは近距離時に" & tmp
  End With
  
  If DDTValue("CVR", 1) > 0 Then
    Cells(14, 6).Value = "検査、調整時はレンズ後方にt=" & _
                         Format(DDTValue("CVR", 1), "0.0#") & _
                         "mm(BK7相当)の平行平面ガラスを挿入すること"
  Else
    Cells(14, 6).Value = "検査、調整時の平行平面ガラスは必要ありません"
  End If
  
  tmp = "&CHOOSE(IF(ISBLANK(F8),1,MID(F8,4,2))"
  For j = 1 To 99
    If DDTValue("EFL", j) <> "NaN" Then
      tmp = tmp & "," & Format(DDTValue("EFL", j), "0.00")
    Else
      Exit For
    End If
  Next
  tmp = tmp & ")&"
  objFormula.SetUp rngCell:=Cells(15, 6), strFormula:= _
    "=" & Chr(34) & "図はf=" & Chr(34) & tmp & Chr(34) & "mmの配置を示す" & Chr(34)
  
  tmp = "諸数値は"
  If DDTValue("REF", 1) > 546 And _
     DDTValue("REF", 1) < 546.2 Then
    tmp = tmp & "e-ray(" & Format(DDTValue("REF", 1), "0.0#") & "nm)"
  ElseIf DDTValue("REF", 1) > 587.5 And _
         DDTValue("REF", 1) < 587.7 Then
    tmp = tmp & "d-ray(" & Format(DDTValue("REF", 1), "0.0#") & "nm)"
  Else
    tmp = tmp & Format(DDTValue("REF", 1), "0.0#") & "nm"
  End If
  Cells(16, 6).Value = tmp & "での値を示す"
Else
  With Cells(7, 8)
    .Value = " NOTE : 赤字の部分は手入力して下さい"
    .Characters(9, 2).Font.ColorIndex = 3
  End With
  With Cells(13, 6)
    .Value = "フォーカスは近距離時に???を？？側に移動させることで行う"
    .Characters(12, 3).Font.ColorIndex = 3
    .Characters(16, 2).Font.ColorIndex = 3
  End With
  With Cells(14, 6)
    .Value = "検査、調整時はレンズ後方にt=?.??mm(BK7相当)の平行平面ガラスを挿入すること"
    .Characters(16, 4).Font.ColorIndex = 3
  End With
  With Cells(15, 6)
    .Value = "図はf=??.??mmの配置を示す"
    .Characters(5, 5).Font.ColorIndex = 3
  End With
  With Cells(16, 6)
    .Value = "諸数値は?-ray(???.??nm)での値を示す"
    .Characters(5, 1).Font.ColorIndex = 3
    .Characters(11, 6).Font.ColorIndex = 3
  End With
End If

With Worksheets(DefaultSheet1)
  n = 1: zRow(1) = 6
  If lngZoo > 0 Then
    For i = lngSur + 9 To lngSur + 8 + lngZoo
      If CLng(Mid(.Cells(i, 4).Value, 2)) > 0 Then
        n = n + 1
        zRow(n) = CLng(Mid(.Cells(i, 4).Value, 2)) + 6
      End If
    Next
  End If
  If zRow(n) < 5 + lngSur Then
    n = n + 1: zRow(n) = 5 + lngSur
  End If
  
  rc(1) = 0: rc(2) = 0: rc(3) = 0
  For i = 1 To n - 1
    chk(1) = False: chk(2) = False
    txt(1) = "": txt(2) = ""
    For j = zRow(i) + 1 To zRow(i + 1)
      If Not IsEmpty(.Cells(j, 3).Value) Then
        If txt(1) = "" Then
          rc(1) = rc(1) + 1
          If rc(1) > 1 Then
            Range(Cells(7 + rc(1), 2), Cells(7 + rc(1), 3)).Insert _
              Shift:=xlShiftDown, CopyOrigin:=xlFormatFromLeftOrAbove
            objBorder.SetUp _
              rngCells:=Range(Cells(7 + rc(1), 2), Cells(7 + rc(1), 3))
          End If
          
          If .Cells(j, 3).Value = "HYB" Then
            If Left(.Cells(j - 1, 3).Value, 1) = "G" Then
              txt(1) = "H" & Mid(.Cells(j - 1, 3).Value, 2)
            ElseIf Left(.Cells(j + 1, 3).Value, 1) = "G" Then
              txt(1) = "H" & Mid(.Cells(j + 1, 3).Value, 2)
            End If
          Else
            txt(1) = .Cells(j, 3).Value
          End If
          txt(2) = txt(1)
          chk(1) = True
          
          If rc(3) > 0 Then eRow(rc(3)) = j - 1
          rc(3) = rc(3) + 1
        Else
          If .Cells(j, 3).Value = "HYB" Then
            If Left(.Cells(j - 1, 3).Value, 1) = "G" Then
              txt(2) = "H" & Mid(.Cells(j - 1, 3).Value, 2)
            ElseIf Left(.Cells(j + 1, 3).Value, 1) = "G" Then
              txt(2) = "H" & Mid(.Cells(j + 1, 3).Value, 2)
            End If
          Else
            txt(2) = .Cells(j, 3).Value
          End If
        End If
        If Left(txt(2), 1) = "G" Then chk(2) = True
        
        If .Cells(j, 3).Value = "STO" Or _
           Left(.Cells(j, 3).Value, 2) = "AP" Then
          sRow(rc(3)) = j
        Else
          sRow(rc(3)) = j + 1
        End If
      End If
    Next
    
    If chk(1) Then
      Cells(7 + rc(1), 2).Value = txt(1) & "-" & txt(2)
      If chk(2) Then
        If DDTValue("EFU", i) <> "NaN" Then
          Cells(7 + rc(1), 3).Value = DDTValue("EFU", i)
        Else
          Cells(7 + rc(1), 3).Value = "??.??"
          Cells(7 + rc(1), 3).Font.ColorIndex = 3
        End If
      Else
        Cells(7 + rc(1), 3).Borders(xlDiagonalUp).LineStyle = xlContinuous
        Cells(7 + rc(1), 3).Borders(xlDiagonalUp).Weight = xlHairline
      End If
      
      txt(3) = "": txt(4) = ""
      For j = zRow(i) + 1 To zRow(i + 1)
        If Not IsEmpty(.Cells(j, 3).Value) Then
          If txt(3) <> "" Then
            If .Cells(j, 3).Value = "STO" Or _
               Left(.Cells(j, 3).Value, 2) = "AP" Then
              txt(4) = .Cells(j, 3).Value
            Else
              If Left(.Cells(j - 1, 3).Value, 1) <> "G" And _
                 .Cells(j - 1, 3).Value <> "HYB" Then
                If .Cells(j, 3).Value = "HYB" Then
                  txt(4) = "H" & Mid(.Cells(j + 1, 3).Value, 2)
                Else
                  txt(4) = .Cells(j, 3).Value
                End If
              End If
            End If
            
            If txt(4) <> "" Then
              rc(2) = rc(2) + 1
              If rc(2) > 1 Then
                Range(Cells(9 + rc(1) + rc(2), 2), Cells(9 + rc(1) + rc(2), 3)).Insert _
                  Shift:=xlShiftDown, CopyOrigin:=xlFormatFromLeftOrAbove
                objBorder.SetUp _
                  rngCells:=Range(Cells(9 + rc(1) + rc(2), 2), Cells(9 + rc(1) + rc(2), 3))
              End If
              Cells(9 + rc(1) + rc(2), 2).Value = txt(3) & "-" & txt(4)
              Cells(9 + rc(1) + rc(2), 3).Value = "+0.05/-0.05"
              txt(3) = "": txt(4) = ""
            End If
          End If
          
          If txt(3) = "" Then
            If .Cells(j, 3).Value = "STO" Or _
               Left(.Cells(j, 3).Value, 2) = "AP" Then
              txt(3) = .Cells(j, 3).Value
            Else
              If Left(.Cells(j + 1, 3).Value, 1) <> "G" And _
                 .Cells(j + 1, 3).Value <> "HYB" Then
                If .Cells(j, 3).Value = "HYB" Then
                  txt(3) = "H" & Mid(.Cells(j - 1, 3).Value, 2)
                Else
                  txt(3) = .Cells(j, 3).Value
                End If
              End If
            End If
          End If
        End If
      Next
    End If
  Next
  
  If DDTValue("TIT", 1) <> "NaN" Then
    If DDTValue("STD", 1) = "Video" Then
      Cells(23, 8).Value = "物体距離 : レンズ前端より"
    ElseIf DDTValue("STD", 1) = "Photo" Then
      Cells(23, 8).Value = "物体距離 : 結像面より"
    End If
    
    sZoo = 1: eZoo = 3: dZoo = 1
    If DDTValue("ZOO", 2) = "NaN" Then
      If DDTValue("ZOO", 3) = "NaN" Then
        eZoo = 1
      Else
        dZoo = 2
      End If
    End If
    
    cc = 0
    For l = sZoo To eZoo Step dZoo
      cc = cc + 1
      pos = DDTValue("ZOO", l)
      Cells(24, 5 + cc).Value = DDTValue("EFL", pos)
      Cells(25, 5 + cc).Value = DDTValue("OLG", l)
      Cells(26, 5 + cc).Value = DDTValue("OLA", l)
      Cells(27, 5 + cc).Value = DDTValue("FNO", l)
      Cells(28, 5 + cc).Value = DDTValue("STO", l) * 2
      Cells(29, 5 + cc).Value = DDTValue("EXP", l)
      Cells(30, 5 + cc).Value = (-1) * DDTValue("MOD", l)
      Cells(31, 5 + cc).Value = (-1) / DDTValue("RED", l)
      Cells(32, 5 + cc).Value = (-1 / 1000) * DDTValue("TST", l)
      
      If rc(3) > 1 Then
        For m = 1 To rc(3) - 1
          If cc = 1 Then
            If m > 1 Then
              Range(Cells(32 + m, 5), Cells(32 + m, 8)).Insert _
                Shift:=xlShiftDown, CopyOrigin:=xlFormatFromLeftOrAbove
              objBorder.SetUp _
                rngCells:=Range(Cells(32 + m, 5), Cells(32 + m, 8))
            End If
            Cells(32 + m, 5).Value = "L" & CStr(m) & "(INF)"
          End If
          
          maxLen = 0
          For i = sRow(m) To eRow(m)
            If IsNumeric(.Cells(i, 6).Value) Then
              maxLen = maxLen + .Cells(i, 6).Value
            Else
              If pos > 0 Then
                tmp = Mid(.Cells(i, 6).Value, 2, Len(.Cells(i, 6).Value) - 2)
                maxRow(1) = .Cells(6, 4).End(xlDown).Row
                maxRow(2) = .Cells(maxRow(1) + 3, 4).End(xlDown).Row
                If maxRow(2) > maxRow(1) + 3 Then
                  For j = maxRow(1) + 4 To maxRow(2)
                    If .Cells(j, 3).Value = tmp Then
                      maxLen = maxLen + .Cells(j, 4 + pos).Value
                      Exit For
                    End If
                  Next
                End If
              End If
            End If
          Next
          Cells(32 + m, 5 + cc).Value = maxLen
        Next
      Else
        If cc = 1 Then Range(Cells(33, 5), Cells(33, 8)).Delete Shift:=xlShiftUp
      End If
      
      Cells(32 + rc(3), 5 + cc).Value = DDTValue("BFG", l)
      Cells(33 + rc(3), 5 + cc).Value = DDTValue("BFA", l)
      Cells(34 + rc(3), 5 + cc).Value = DDTValue("FCS", l)
      Cells(35 + rc(3), 5 + cc).Value = DDTValue("FCM", l)
    Next
  Else
    Cells(23, 8).Value = "物体距離 : ？？？より"
    Cells(23, 8).Characters(8, 3).Font.ColorIndex = 3
    
    sZoo = 1: eZoo = 3: dZoo = 1
    cc = 0
    For l = sZoo To eZoo Step dZoo
      cc = cc + 1
      Cells(24, 5 + cc).Value = "??.??"
      Cells(25, 5 + cc).Value = "??.??"
      Cells(26, 5 + cc).Value = "??.??"
      Cells(27, 5 + cc).Value = "??.??"
      Cells(28, 5 + cc).Value = "??.??"
      Cells(29, 5 + cc).Value = "??.??"
      Cells(30, 5 + cc).Value = "??.??"
      Cells(31, 5 + cc).Value = "??.??"
      Cells(32, 5 + cc).Value = "??.??"
      
      If rc(3) > 1 Then
        For m = 1 To rc(3) - 1
          If cc = 1 Then
            If m > 1 Then
              Range(Cells(32 + m, 5), Cells(32 + m, 8)).Insert _
                Shift:=xlShiftDown, CopyOrigin:=xlFormatFromLeftOrAbove
              objBorder.SetUp _
                rngCells:=Range(Cells(32 + m, 5), Cells(32 + m, 8))
            End If
            Cells(32 + m, 5).Value = "L" & CStr(m)
          End If
          
          Cells(32 + m, 5 + cc).Value = "??.??"
        Next
      Else
        If cc = 1 Then Range(Cells(33, 5), Cells(33, 8)).Delete Shift:=xlShiftUp
      End If
      
      Cells(32 + rc(3), 5 + cc).Value = "??.??"
      Cells(33 + rc(3), 5 + cc).Value = "??.??"
      Cells(34 + rc(3), 5 + cc).Value = "??.??"
      Cells(35 + rc(3), 5 + cc).Value = "??.??"
    Next
    
    Range(Cells(24, 6), Cells(35 + rc(3), 8)).Font.ColorIndex = 3
  End If
End With

Set objFormula = Nothing
Set objBorder = Nothing

End Sub


Public Sub AddG()

Dim xlWs As Excel.Worksheet
Dim fsObject As Object
Dim objFormula As clsFormula

Dim f As Integer
Dim i As Long, rc As Long, af As Long, co As Long
Dim tmp As Double, h(2) As Double
Dim gcFileName As String, TextLine As String
Dim fml(2) As String, al(2) As String
Dim strGla(0 To 4) As String

Set objFormula = New clsFormula

fml(1) = ",C19,C20,C21,C22,C23,C24,C25"
fml(2) = ",C26,C27,C28,C29,C30,C31,C32"

Cells(2, 3).Value = mdlNum
Cells(3, 3).Value = "G" & Mid(ActiveSheet.Name, 2, 2)
Cells(4, 3).Value = Date

With Worksheets(DefaultSheet1)
  rc = 0
  If lngSur > 3 Then
    For i = 7 To lngSur + 3
      If .Cells(i, 3).Value = "G" & Mid(ActiveSheet.Name, 2, 2) Then
        rc = i: Exit For
      End If
    Next
  End If
  
  If rc > 0 Then
    If intOut <> 0 Then
      rc = rc + (intOut * 2 - 3)
    End If
    
    Cells(8, 3).Value = .Cells(rc, 5).Value
    Cells(9, 3).Value = .Cells(rc + 1, 5).Value
    Cells(10, 3).Value = .Cells(rc, 6).Value
    If intOut <> 2 Then
      Cells(13, 3).Value = .Cells(rc, 9).Value * 2
    End If
    If intOut <> 1 Then
      Cells(14, 3).Value = .Cells(rc + 1, 9).Value * 2
    End If
    
    Cells(11, 3).Value = .Cells(rc, 7).Value
    Cells(12, 3).Value = .Cells(rc, 8).Value
    Cells(8, 9).Value = 1
    Cells(9, 9).Value = 0
    Cells(16, 9).Value = 0
    Set fsObject = CreateObject("Scripting.FileSystemObject")
    gcFileName = _
      ThisWorkbook.Path & "\GlassCatalog\" & .Cells(rc, 8).Value & ".csv"
    If fsObject.FileExists(gcFileName) Then
      f = FreeFile(0)
      Open gcFileName For Input As #f
      Do While Not EOF(f)
        Line Input #f, TextLine
        If InStr(TextLine, ",") > 0 Then
          strGla(0) = UCase(Mid(TextLine, 1, InStr(TextLine, ",") - 1))
          If .Cells(rc, 8).Value <> "OTHER" Then
            strGla(0) = Replace(strGla(0), "-", "")
            strGla(0) = Replace(strGla(0), " ", "")
          End If
          If strGla(0) = .Cells(rc, 7).Value Then
            TextLine = TextLine & ",": co = 0
            For i = 1 To 4
              strGla(i) = Trim(Mid(TextLine, co + 1, InStr(co + 1, TextLine, ",") - (co + 1)))
              co = InStr(co + 1, TextLine, ",")
            Next
            .Cells(rc, 7).Value = strGla(1)
            Cells(11, 3).Value = strGla(1)
            Cells(8, 9).Value = strGla(2)
            Cells(9, 9).Value = strGla(3)
            Cells(16, 9).Value = strGla(4)
            Exit Do
          End If
        End If
      Loop
      Close #f
    End If
    Set fsObject = Nothing
    
    If intOut = 0 Then
      objFormula.SetUp _
        rngCell:=Cells(14, 9), strFormula:="=FocalLength(C8,C9,C10,I8)"
    End If
    
    af = 1
    If InStr(.Cells(rc, 4).Value, "A") = 0 Then
      al(1) = ""
    Else
      al(1) = Mid(.Cells(rc, 4).Value, InStr(.Cells(rc, 4).Value, "A"), _
              Len(.Cells(rc, 4).Value) - InStr(.Cells(rc, 4).Value, "A"))
      af = af + 1
    End If
    If InStr(.Cells(rc + 1, 4).Value, "A") = 0 Then
      al(2) = ""
    Else
      al(2) = Mid(.Cells(rc + 1, 4).Value, InStr(.Cells(rc + 1, 4).Value, "A"), _
              Len(.Cells(rc + 1, 4).Value) - InStr(.Cells(rc + 1, 4).Value, "A"))
      af = af + 1
    End If
    If lngAsp > 0 Then
      For i = lngSur + lngZoo + 12 To lngSur + lngZoo + 11 + lngAsp
        Select Case .Cells(i, 3).Value
          Case al(1)
            Cells(8, 3).NumberFormat = "+0.0000;-0.0000"
            Cells(19, 3).Value = .Cells(i, 6).Value
            Cells(20, 3).Value = .Cells(i, 7).Value
            Cells(21, 3).Value = .Cells(i, 8).Value
            Cells(22, 3).Value = .Cells(i, 9).Value
            Cells(23, 3).Value = .Cells(i, 10).Value
            Cells(24, 3).Value = .Cells(i, 11).Value
            Cells(25, 3).Value = .Cells(i, 12).Value
          Case al(2)
            Cells(9, 3).NumberFormat = "+0.0000;-0.0000"
            Cells(26, 3).Value = .Cells(i, 6).Value
            Cells(27, 3).Value = .Cells(i, 7).Value
            Cells(28, 3).Value = .Cells(i, 8).Value
            Cells(29, 3).Value = .Cells(i, 9).Value
            Cells(30, 3).Value = .Cells(i, 10).Value
            Cells(31, 3).Value = .Cells(i, 11).Value
            Cells(32, 3).Value = .Cells(i, 12).Value
        End Select
      Next
    End If
    If af > 1 Then ActiveSheet.Name = ActiveSheet.Name & "(ASP)"
    
    Select Case Left(ActiveSheet.Name, 1)
      Case "G"
        h(1) = CDbl(Alpha(1)): h(2) = CDbl(Alpha(1))
        If Cells(14, 9).Value > 0 Then
          Select Case af
            Case 2: h(1) = CDbl(Alpha(6)): h(2) = CDbl(Alpha(6))
            Case 3: h(1) = CDbl(Alpha(8)): h(2) = CDbl(Alpha(8))
          End Select
        Else
          If IsNumeric(.Cells(rc, 5).Value) Then
            If .Cells(rc, 5).Value < 0 Then
              h(1) = CDbl(Alpha(2)) + CDbl(Alpha(4))
            End If
          End If
          If IsNumeric(.Cells(rc + 1, 5).Value) Then
            If .Cells(rc + 1, 5).Value > 0 Then
              h(2) = CDbl(Alpha(2)) + CDbl(Alpha(4))
            End If
          End If
        End If
        h(1) = h(1) + .Cells(rc, 9).Value: h(2) = h(2) + .Cells(rc + 1, 9).Value
        
        If h(1) >= h(2) Then
          Cells(8, 6).Value = h(1) * 2
        Else
          Cells(8, 6).Value = h(2) * 2
        End If
        
        Cells(9, 6).Value = Cells(8, 6).Value: Cells(10, 6).Value = Cells(8, 6).Value
        If Cells(14, 9).Value > 0 Then
          If af = 3 Then
            Cells(8, 6).Value = Cells(8, 6).Value + CDbl(Alpha(9)) * 2
            Cells(9, 6).Value = h(1) * 2: Cells(10, 6).Value = h(2) * 2
          End If
        Else
          If IsNumeric(.Cells(rc, 5).Value) Then
            If .Cells(rc, 5).Value < 0 Then
              Cells(9, 6).Value = (h(1) - CDbl(Alpha(4))) * 2
            End If
          End If
          If IsNumeric(.Cells(rc + 1, 5).Value) Then
            If .Cells(rc + 1, 5).Value > 0 Then
              Cells(10, 6).Value = (h(2) - CDbl(Alpha(4))) * 2
            End If
          End If
        End If
        
        objFormula.SetUp _
          rngCell:=Cells(11, 6), strFormula:="=ABS(Sag(C8,F9" & fml(1) & "))"
        objFormula.SetUp _
          rngCell:=Cells(12, 6), strFormula:="=ABS(Sag(C9,F10" & fml(2) & "))"
        Cells(13, 6).Value = 0.2
        Cells(14, 6).Value = 0.2
        If Cells(14, 9).Value > 0 Then
          objFormula.SetUp rngCell:=Cells(15, 6), strFormula:= _
            "=C10-F13-F14-Sag(C8,F9-IF(F8>F9,0,F13)*2" & fml(1) & ")" & _
                        "+Sag(C9,F10-IF(F8>F10,0,F14)*2" & fml(2) & ")"
          Select Case af
            Case 1
              If Alpha(3) = "A" Then
                objFormula.SetUp rngCell:=Cells(16, 6), strFormula:= _
                  "=C10-Sag(C8,F8+PressValue(C8,C9,C13,C14,F8)" & fml(1) & ")" & _
                      "+Sag(C9,F8+PressValue(C8,C9,C13,C14,F8)" & fml(2) & ")"
              Else
                objFormula.SetUp rngCell:=Cells(16, 6), strFormula:= _
                  "=C10-Sag(C8,F8+" & Alpha(3) & "*2" & fml(1) & ")" & _
                      "+Sag(C9,F8+" & Alpha(3) & "*2" & fml(2) & ")"
              End If
            Case 2
              objFormula.SetUp rngCell:=Cells(16, 6), strFormula:= _
                "=C10-Sag(C8,F8+" & Alpha(7) & "*2" & fml(1) & ")" & _
                    "+Sag(C9,F8+" & Alpha(7) & "*2" & fml(2) & ")"
          End Select
        End If
        
        If Left(.Cells(rc - 1, 3).Value, 1) = "G" Then
          Cells(10, 9).Value = "-"
        Else
          Cells(10, 9).Value = "540±20nm"
        End If
        If Left(.Cells(rc + 1, 3).Value, 1) = "G" Then
          Cells(11, 9).Value = "-"
        Else
          Cells(11, 9).Value = "540±20nm"
        End If
        objFormula.SetUp rngCell:=Cells(12, 9), strFormula:= _
          "=IF(AND(F8>F9,C8<0,C8<>""Inf""),""研磨面全面"",F8-" & Alpha(5) & "*2)"
        objFormula.SetUp rngCell:=Cells(13, 9), strFormula:= _
          "=IF(AND(F8>F10,C9>0,C9<>""Inf""),""研磨面全面"",F8-" & Alpha(5) & "*2)"
        objFormula.SetUp rngCell:=Cells(15, 9), strFormula:= _
          "=GlassWeight(C8,C9,C10,I16,F9,F10,F8" & fml(1) & fml(2) & ")"
        
        If af > 1 Then
          Range(Cells(21, 6), Cells(23, 9)).Borders(xlDiagonalUp).LineStyle = xlContinuous
          Range(Cells(21, 6), Cells(23, 9)).Borders(xlDiagonalUp).Weight = xlHairline
        End If
        
      Case "H"
        Set xlWs = Worksheets("G" & Mid(ActiveSheet.Name, 2, 2))
        
        Cells(3, 3).Value = Cells(3, 3).Value & " " & Alpha(15)
        xlWs.Cells(3, 3).Value = xlWs.Cells(3, 3).Value & " " & Alpha(16)
        
        Select Case intOut
          Case 1
            tmp = .Cells(rc, 9).Value * 2
            Cells(8, 6).Value = tmp + 0.7 * 2
            Cells(9, 6).Value = tmp + 0.5 * 2
            Cells(10, 6).Value = tmp + 0.7 * 2
            If xlWs.Cells(8, 6).Value > xlWs.Cells(9, 6).Value Then
              Cells(10, 6).Value = xlWs.Cells(9, 6).Value
            End If
            Cells(10, 9).Value = "AM3"
            Cells(11, 9).Value = "-"
            Cells(12, 9).Value = tmp + 0.5 * 2
          Case 2
            tmp = .Cells(rc + 1, 9).Value * 2
            Cells(8, 6).Value = tmp + 0.7 * 2
            Cells(9, 6).Value = tmp + 0.7 * 2
            Cells(10, 6).Value = tmp + 0.5 * 2
            If xlWs.Cells(8, 6).Value > xlWs.Cells(10, 6).Value Then
              Cells(9, 6).Value = xlWs.Cells(10, 6).Value
            End If
            Cells(10, 9).Value = "-"
            Cells(11, 9).Value = "AM3"
            Cells(13, 9).Value = tmp + 0.5 * 2
        End Select
        
        If xlWs.Cells(8, 6).Value < tmp + 1.2 * 2 Then
          If xlWs.Cells(8, 6).Value = xlWs.Cells(9, 6).Value Then
            xlWs.Cells(9, 6).Value = tmp + 1.2 * 2
          End If
          If xlWs.Cells(8, 6).Value = xlWs.Cells(10, 6).Value Then
            xlWs.Cells(10, 6).Value = tmp + 1.2 * 2
          End If
          xlWs.Cells(8, 6).Value = tmp + 1.2 * 2
        End If
        
        Set xlWs = Nothing
    End Select
    
    intOut = 0
    If .Cells(rc - 1, 3).Value = "HYB" Then intOut = 1
    If .Cells(rc + 1, 3).Value = "HYB" Then intOut = 2
  End If
End With

Set objFormula = Nothing

End Sub


Public Sub AddB()

Dim xlWs1 As Excel.Worksheet
Dim xlWs2 As Excel.Worksheet
Dim Unused As Range

Dim i As Long, rc As Long
Dim tmp As Double
Dim txt(2) As String

Cells(8, 2).Value = "構成部番1"
Cells(9, 2).Value = "構成部番2"
Cells(10, 2).Value = "光学有効径1"
Cells(11, 2).Value = "光学有効径2"
Cells(12, 2).Value = "光学有効径3"
Cells(18, 2).Value = "■合成肉厚公差"
Cells(19, 2).Value = ""

With Range(Cells(8, 3), Cells(9, 3))
  .NumberFormat = "General"
  .HorizontalAlignment = xlRight
End With
Range(Cells(10, 3), Cells(12, 3)).NumberFormat = """φ""0.00"
Range(Cells(11, 3), Cells(12, 3)).HorizontalAlignment = xlGeneral
With Cells(19, 3)
  .ShrinkToFit = True
  .NumberFormat = "@"
  .HorizontalAlignment = xlRight
End With

Set Unused = Union(Range(Cells(13, 2), Cells(14, 3)), _
                   Range(Cells(20, 2), Cells(32, 3)), _
                   Range(Cells(7, 5), Cells(16, 6)), _
                   Range(Cells(30, 5), Cells(36, 6)), _
                   Range(Cells(7, 8), Cells(16, 9)))
With Unused
  .Clear
  .Font.Name = "ＭＳ 明朝"
  .Font.Size = 10
End With
Set Unused = Nothing

With Range(Cells(18, 2), Cells(19, 3))
  .Cut Destination:=Range(Cells(14, 2), Cells(15, 3))
  .Font.Name = "ＭＳ 明朝"
  .Font.Size = 10
End With
With Range(Cells(20, 5), Cells(28, 9))
  .Cut Destination:=Range(Cells(7, 5), Cells(15, 9))
  .Font.Name = "ＭＳ 明朝"
  .Font.Size = 10
End With

Cells(2, 3).Value = mdlNum
Cells(3, 3).Value = ActiveSheet.Name
Cells(4, 3).Value = Date

With Worksheets(DefaultSheet1)
  rc = 0
  If lngSur > 3 Then
    For i = 7 To lngSur + 3
      If .Cells(i, 2).Value = ActiveSheet.Name Then
        rc = i: Exit For
      End If
    Next
  End If
  
  If rc > 0 Then
    txt(1) = .Cells(rc, 3).Value
    txt(2) = .Cells(rc + 1, 3).Value
    If InStr(.Cells(rc, 4).Value, "A") > 0 Or _
       InStr(.Cells(rc + 1, 4).Value, "A") > 0 Then
      Set xlWs1 = Worksheets(txt(1) & "(ASP)")
    Else
      Set xlWs1 = Worksheets(txt(1))
    End If
    If InStr(.Cells(rc + 1, 4).Value, "A") > 0 Or _
       InStr(.Cells(rc + 2, 4).Value, "A") > 0 Then
      Set xlWs2 = Worksheets(txt(2) & "(ASP)")
    Else
      Set xlWs2 = Worksheets(txt(2))
    End If
    Cells(3, 3).Value = Cells(3, 3).Value & "(" & txt(1) & _
                                            "/" & txt(2) & "接合図)"
    
    If .Cells(rc - 1, 3).Value = "HYB" Then
      Cells(8, 3).Value = "M" & Mid(xlWs1.Name, 2, 2)
      Cells(10, 3).Value = .Cells(rc - 1, 9).Value * 2
    Else
      Cells(8, 3).Value = xlWs1.Name
      Cells(10, 3).Value = .Cells(rc, 9).Value * 2
    End If
    If .Cells(rc + 2, 3).Value = "HYB" Then
      Cells(9, 3).Value = "M" & Mid(xlWs2.Name, 2, 2)
      Cells(12, 3).Value = .Cells(rc + 3, 9).Value * 2
    Else
      Cells(9, 3).Value = xlWs2.Name
      Cells(12, 3).Value = .Cells(rc + 2, 9).Value * 2
    End If
    Cells(11, 3).Value = .Cells(rc + 1, 9).Value * 2
    Cells(15, 2).Value = Left(Cells(8, 3).Value, 3) & "+" & Left(Cells(9, 3).Value, 3)
    Cells(15, 3).Value = "+[累積]/-[累積]"
    
    Cells(8, 6).Value = "光学有効径 R1:" & Cells(10, 3).Text & _
                        ", 接合面:" & Cells(11, 3).Text & _
                        ", R2:" & Cells(12, 3).Text
    
    Cells(9, 6).Value = "合成芯精度:"
    If xlWs1.Cells(14, 9).Value > 0 Then
      Cells(9, 6).Value = Cells(9, 6).Value & txt(2)
    Else
      Cells(9, 6).Value = Cells(9, 6).Value & txt(1)
    End If
    Cells(9, 6).Value = Cells(9, 6).Value & "基準にて透過芯1'以内"
    
    If xlWs1.Cells(14, 9).Value > 0 Then
      If xlWs2.Cells(8, 3).Value < 0 Then
        tmp = (xlWs1.Cells(10, 6).Value - _
               xlWs1.Cells(14, 6).Value * 2) - xlWs2.Cells(9, 6).Value
        If tmp > 0 Then
          If xlWs2.Cells(8, 6).Value = xlWs2.Cells(10, 6).Value Then
            xlWs2.Cells(10, 6).Value = xlWs2.Cells(10, 6).Value + tmp
          End If
          xlWs2.Cells(8, 6).Value = xlWs2.Cells(8, 6).Value + tmp
          xlWs2.Cells(9, 6).Value = xlWs2.Cells(9, 6).Value + tmp
        End If
      End If
    Else
      If xlWs1.Cells(9, 3).Value > 0 Then
        tmp = (xlWs2.Cells(9, 6).Value - _
               xlWs2.Cells(13, 6).Value * 2) - xlWs1.Cells(10, 6).Value
        If tmp > 0 Then
          If xlWs1.Cells(8, 6).Value = xlWs1.Cells(9, 6).Value Then
            xlWs1.Cells(9, 6).Value = xlWs1.Cells(9, 6).Value + tmp
          End If
          xlWs1.Cells(8, 6).Value = xlWs1.Cells(8, 6).Value + tmp
          xlWs1.Cells(10, 6).Value = xlWs1.Cells(10, 6).Value + tmp
        End If
      End If
    End If
    Set xlWs1 = Nothing
    Set xlWs2 = Nothing
  End If
End With

End Sub


Public Sub AddM()

Dim xlWs1 As Excel.Worksheet
Dim xlWs2 As Excel.Worksheet
Dim Unused As Range

Cells(8, 2).Value = "構成部番1"
Cells(9, 2).Value = "構成部番2"
Cells(10, 2).Value = "光学有効径1"
Cells(11, 2).Value = "光学有効径2"

With Range(Cells(8, 3), Cells(9, 3))
  .NumberFormat = "General"
  .HorizontalAlignment = xlRight
End With
Range(Cells(10, 3), Cells(11, 3)).NumberFormat = """φ""0.00"
Cells(11, 3).HorizontalAlignment = xlGeneral

Set Unused = Union(Range(Cells(12, 2), Cells(14, 3)), _
                   Range(Cells(18, 2), Cells(32, 3)), _
                   Range(Cells(7, 5), Cells(16, 6)), _
                   Range(Cells(30, 5), Cells(36, 6)), _
                   Range(Cells(7, 8), Cells(16, 9)))
With Unused
  .Clear
  .Font.Name = "ＭＳ 明朝"
  .Font.Size = 10
End With
Set Unused = Nothing

With Range(Cells(20, 5), Cells(28, 9))
  .Cut Destination:=Range(Cells(7, 5), Cells(15, 9))
  .Font.Name = "ＭＳ 明朝"
  .Font.Size = 10
End With

Cells(2, 3).Value = mdlNum
Cells(3, 3).Value = "G" & Mid(ActiveSheet.Name, 2, 2)
Cells(4, 3).Value = Date

Set xlWs1 = Worksheets("G" & Mid(ActiveSheet.Name, 2, 2))
Set xlWs2 = Worksheets("H" & Mid(ActiveSheet.Name, 2, 2) & "(ASP)")
If Not IsEmpty(xlWs2.Cells(19, 3).Value) Then
  Cells(8, 3).Value = xlWs2.Name
  Cells(9, 3).Value = xlWs1.Name
ElseIf Not IsEmpty(xlWs2.Cells(26, 3).Value) Then
  Cells(8, 3).Value = xlWs1.Name
  Cells(9, 3).Value = xlWs2.Name
End If
Cells(10, 3).Value = Worksheets(Cells(8, 3).Value).Cells(13, 3).Value
Cells(11, 3).Value = Worksheets(Cells(9, 3).Value).Cells(14, 3).Value

Cells(8, 6).Value = "光学有効径 R1:" & Cells(10, 3).Text & _
                             ", R2:" & Cells(11, 3).Text
Cells(9, 6).Value = "合成芯精度:" & Cells(3, 3).Value & _
                    Alpha(16) & "基準にて透過芯1.5'以内"

Set xlWs1 = Nothing
Set xlWs2 = Nothing

End Sub


Private Function DDTValue(ByVal strCommand As String, ByVal lngIndex As Long) As Variant

Dim i As Long, j As Long
Dim sp As Long

DDTValue = "NaN"
For i = 1 To UBound(DDT)
  If DDT(i) = "" Then Exit Function
  If Left(DDT(i), 3) = strCommand Then
    sp = 5
    For j = 1 To 99
      If InStr(sp + 1, DDT(i), " ") > 0 Then
        If j = lngIndex Then
          DDTValue = Trim(Mid(DDT(i), sp + 1, InStr(sp + 1, DDT(i), " ") - (sp + 1)))
          Exit For
        End If
        sp = InStr(sp + 1, DDT(i), " ")
      Else
        Exit For
      End If
    Next
    Exit For
  End If
Next

End Function


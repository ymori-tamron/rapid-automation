VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsXltxFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False


Option Explicit
Option Base 1
Option Compare Binary


Public Sub DataSummary()

Dim objBorder As clsBorder

Dim i As Long, snw As Long

Set objBorder = New clsBorder

With Application
  snw = .SheetsInNewWorkbook
  .SheetsInNewWorkbook = 1
  Workbooks.Add
  .SheetsInNewWorkbook = snw
End With
With Cells
  .ClearFormats
  .Font.Name = "ＭＳ 明朝"
  .Font.Size = 10
End With
ActiveWindow.Zoom = 75
ActiveSheet.Name = DefaultSheet1

Columns(1).ColumnWidth = 2
Range(Columns(2), Columns(3)).ColumnWidth = 4.5
Columns(4).ColumnWidth = 8.5
Range(Columns(5), Columns(12)).ColumnWidth = 14
Columns(13).ColumnWidth = 2


Worksheets.Add After:=Worksheets(Worksheets.Count)
With Cells
  .ClearFormats
  .Font.Name = "ＭＳ 明朝"
  .Font.Size = 10
End With
ActiveWindow.Zoom = 100
ActiveSheet.Name = DefaultSheet2

For i = 1 To 7 Step 3
  Columns(i).ColumnWidth = 2
  Range(Columns(i + 1), Columns(i + 2)).ColumnWidth = 14
Next
Columns(10).ColumnWidth = 2

Cells(2, 2).Value = "図面番号"
Cells(3, 2).Value = "部品番号"
Cells(4, 2).Value = "作成日"
Cells(5, 2).Value = "品目コード"
Cells(2, 5).Value = "尺度"

Cells(7, 2).Value = "■基本データ"
Cells(8, 2).Value = "曲率半径1"
Cells(9, 2).Value = "曲率半径2"
Cells(10, 2).Value = "肉厚"
Cells(11, 2).Value = "硝種"
Cells(12, 2).Value = "メーカー名"
Cells(13, 2).Value = "光学有効径1"
Cells(14, 2).Value = "光学有効径2"

Cells(18, 2).Value = "■非球面データ(非球面式はCODEV形式)"
Cells(19, 2).Value = "コーニック定数1"
Cells(20, 2).Value = "4次の係数1"
Cells(21, 2).Value = "6次の係数1"
Cells(22, 2).Value = "8次の係数1"
Cells(23, 2).Value = "10次の係数1"
Cells(24, 2).Value = "12次の係数1"
Cells(25, 2).Value = "14次の係数1"
Cells(26, 2).Value = "コーニック定数2"
Cells(27, 2).Value = "4次の係数2"
Cells(28, 2).Value = "6次の係数2"
Cells(29, 2).Value = "8次の係数2"
Cells(30, 2).Value = "10次の係数2"
Cells(31, 2).Value = "12次の係数2"
Cells(32, 2).Value = "14次の係数2"

Cells(7, 5).Value = "■形状データ"
Cells(8, 5).Value = "外径"
Cells(9, 5).Value = "内径1"
Cells(10, 5).Value = "内径2"
Cells(11, 5).Value = "⊿Ｈ1"
Cells(12, 5).Value = "⊿Ｈ2"
Cells(13, 5).Value = "面取り1"
Cells(14, 5).Value = "面取り2"
Cells(15, 5).Value = "芯取上コバ厚"
Cells(16, 5).Value = "研磨上コバ厚"

Cells(7, 8).Value = "■記載データ"
Cells(8, 8).Value = "Ｎd"
Cells(9, 8).Value = "νd"
Cells(10, 8).Value = "コート1"
Cells(11, 8).Value = "コート2"
Cells(12, 8).Value = "コート有効径1"
Cells(13, 8).Value = "コート有効径2"
Cells(14, 8).Value = "焦点距離(fd)"
Cells(15, 8).Value = "重量"
Cells(16, 8).Value = "比重"

Cells(20, 5).Value = "■注記データ"
Cells(21, 5).Value = "注記1"
Cells(22, 5).Value = "注記2"
Cells(23, 5).Value = "注記3"
Cells(24, 5).Value = "注記4"
Cells(25, 5).Value = "注記5"
Cells(26, 5).Value = "注記6"
Cells(27, 5).Value = "注記7"
Cells(28, 5).Value = "注記8"

Cells(30, 5).Value = "■類似硝種データ"
Cells(31, 5).Value = "HOYA"
Cells(32, 5).Value = "OHARA"
Cells(33, 5).Value = "SCHOTT"
Cells(34, 5).Value = "光ガラス"
Cells(35, 5).Value = "成都光明"
Cells(36, 5).Value = "新華光"

With Cells(2, 6)
  With .Validation
    .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
         Formula1:="自動,1倍図,2倍図,5倍図,10倍図"
    .InCellDropdown = True
    .IgnoreBlank = False
    .ShowInput = False
    .ShowError = True
    .ErrorTitle = ProjectTitle
    .ErrorMessage = "リスト以外の項目は入力出来ません"
    .IMEMode = xlIMEModeNoControl
  End With
  .Value = "自動"
End With

Range(Cells(8, 3), Cells(9, 3)).NumberFormat = "+0.0##;-0.0##"
Cells(10, 3).NumberFormat = "0.0#"
Range(Cells(11, 3), Cells(12, 3)).HorizontalAlignment = xlRight
Range(Cells(13, 3), Cells(14, 3)).NumberFormat = """φ""0.00"
objBorder.SetUp rngCells:=Range(Cells(8, 2), Cells(14, 3))

Cells(19, 2).ShrinkToFit = True
Cells(19, 3).NumberFormat = "+0.0000;-0.0000"
Range(Cells(20, 3), Cells(25, 3)).NumberFormat = "+0.00000E+00;-0.00000E+00"
Cells(26, 2).ShrinkToFit = True
Cells(26, 3).NumberFormat = "+0.0000;-0.0000"
Range(Cells(27, 3), Cells(32, 3)).NumberFormat = "+0.00000E+00;-0.00000E+00"
objBorder.SetUp rngCells:=Range(Cells(19, 2), Cells(32, 3))

Cells(8, 6).NumberFormat = """φ""0.00"
Range(Cells(9, 6), Cells(10, 6)).NumberFormat = """(φ""0.00"")"""
Range(Cells(11, 6), Cells(12, 6)).NumberFormat = "0.000"
Range(Cells(13, 6), Cells(14, 6)).NumberFormat = """C""0.0#"
Range(Cells(15, 6), Cells(16, 6)).NumberFormat = "0.000"
objBorder.SetUp rngCells:=Range(Cells(8, 5), Cells(16, 6))

Cells(8, 9).NumberFormat = "0.00000"
Cells(9, 9).NumberFormat = "0.0#"
Range(Cells(10, 9), Cells(13, 9)).HorizontalAlignment = xlRight
Range(Cells(12, 9), Cells(13, 9)).NumberFormat = """φ""0.00"
Cells(14, 9).NumberFormat = "+0.000;-0.000"
Cells(15, 9).NumberFormat = "0.000""g"""
objBorder.SetUp rngCells:=Range(Cells(8, 8), Cells(16, 9))

With Range(Cells(21, 6), Cells(28, 9))
  .ShrinkToFit = True
  .Merge Across:=True
End With
objBorder.SetUp rngCells:=Range(Cells(21, 5), Cells(28, 9))

Range(Cells(31, 6), Cells(36, 6)).HorizontalAlignment = xlRight
objBorder.SetUp rngCells:=Range(Cells(31, 5), Cells(36, 6))


Worksheets.Add After:=Worksheets(Worksheets.Count)
With Cells
  .ClearFormats
  .Font.Name = "ＭＳ 明朝"
  .Font.Size = 10
End With
With ActiveWindow
  .Zoom = 85
  .SplitRow = 6
  .SplitColumn = 2
  .FreezePanes = True
End With
ActiveSheet.Name = DefaultSheet3

Columns(1).ColumnWidth = 2
Columns(2).ColumnWidth = 14
Range(Columns(3), Columns(4)).ColumnWidth = 8
Columns(5).ColumnWidth = 2

Cells(2, 2).Value = "図面番号"
Cells(3, 2).Value = "作成日"

Cells(5, 2).Value = "■公差データ"
Cells(7, 2).Value = "ニュートン1"
Cells(8, 2).Value = "ニュートン2"
Cells(9, 2).Value = "アス1"
Cells(10, 2).Value = "アス2"
Cells(11, 2).Value = "肉厚"
Cells(12, 2).Value = "Ｎd"
Cells(13, 2).Value = "νd"
Cells(14, 2).Value = "外径"
Cells(15, 2).Value = "内径1"
Cells(16, 2).Value = "内径2"
Cells(17, 2).Value = "⊿Ｈ1"
Cells(18, 2).Value = "⊿Ｈ2"
Cells(19, 2).Value = "面取り1"
Cells(20, 2).Value = "面取り2"
Cells(21, 2).Value = "透過芯"
Cells(22, 2).Value = "反射芯"
Cells(23, 2).Value = "垂直度"
Cells(24, 2).Value = "平行度"
Cells(25, 2).Value = "偏芯面"
Cells(26, 2).Value = "型保証偏芯"

Range(Cells(2, 3), Cells(3, 4)).Merge Across:=True
With Range(Cells(6, 3), Cells(6, 4))
  .Merge
  .HorizontalAlignment = xlCenter
End With
Range(Cells(7, 3), Cells(10, 3)).NumberFormat = _
  """±""General""本"";""±""General""本"""
Range(Cells(7, 4), Cells(10, 4)).NumberFormat = "0.00""μ"""
Range(Cells(11, 3), Cells(11, 4)).NumberFormat = "+0.0##;-0.0##"
Cells(12, 3).NumberFormat = """±""0.0000"
Cells(12, 4).NumberFormat = """±""General""E-5"""
Cells(13, 3).NumberFormat = """±""0.0%"
Cells(13, 4).NumberFormat = """±""0.0#"
Range(Cells(14, 3), Cells(20, 4)).NumberFormat = "+0.0##;-0.0##"
Range(Cells(14, 3), Cells(16, 3)).HorizontalAlignment = xlRight
Range(Cells(19, 3), Cells(20, 3)).HorizontalAlignment = xlRight
Range(Cells(21, 3), Cells(24, 3)).NumberFormat = "General""'"""
Range(Cells(22, 4), Cells(24, 4)).NumberFormat = "0.000"
Cells(25, 3).NumberFormat = """R""0"
With Range(Cells(25, 3), Cells(25, 4))
  .Merge
  .HorizontalAlignment = xlCenter
End With
Cells(26, 3).NumberFormat = "General""μ"""
Cells(26, 4).NumberFormat = "General""'"""
Range(Cells(6, 3), Cells(26, 4)).ShrinkToFit = True
objBorder.SetUp rngCells:=Range(Cells(6, 2), Cells(26, 4))

Set objBorder = Nothing

End Sub


Public Sub SagTable()

Dim i As Long, snw As Long

With Application
  snw = .SheetsInNewWorkbook
  .SheetsInNewWorkbook = 1
  Workbooks.Add
  .SheetsInNewWorkbook = snw
End With
With Cells
  .ClearFormats
  .Font.Name = "ＭＳ 明朝"
  .Font.Size = 10
End With
ActiveWindow.Zoom = 75
With ActiveSheet
  .Name = "非球面形状"
  With .PageSetup
    .Orientation = xlPortrait
    .Zoom = 95
    .PaperSize = xlPaperA4
    .PrintQuality = 600
    .FirstPageNumber = xlAutomatic
    .LeftMargin = Application.CentimetersToPoints(2)
    .RightMargin = Application.CentimetersToPoints(1)
    .TopMargin = Application.CentimetersToPoints(1)
    .BottomMargin = Application.CentimetersToPoints(1)
    .HeaderMargin = Application.CentimetersToPoints(0)
    .FooterMargin = Application.CentimetersToPoints(0)
    .CenterHorizontally = False
    .CenterVertically = False
  End With
  .HPageBreaks.Add Before:=Cells(73, 1)
  .VPageBreaks.Add Before:=Cells(1, 14)
End With

Columns(1).ColumnWidth = 2
For i = 2 To 11 Step 3
  Columns(i).ColumnWidth = 6
  Columns(i + 1).ColumnWidth = 14
  Columns(i + 2).ColumnWidth = 2
Next

Cells(6, 2).Value = "Ｒ"
Cells(7, 2).Value = "ε"
Cells(8, 2).Value = "Ａ"
Cells(9, 2).Value = "Ｂ"
Cells(10, 2).Value = "Ｃ"
Cells(11, 2).Value = "Ｄ"
Cells(12, 2).Value = "Ｅ"
Cells(13, 2).Value = "Ｆ"
Cells(14, 2).Value = "Ｇ"

Cells(16, 2).Value = "H"
Cells(16, 3).Value = "X"

With Cells(2, 2).Font
  .Size = 12
  .FontStyle = "Italic"
  .Underline = xlUnderlineStyleSingle
End With
Cells(3, 3).HorizontalAlignment = xlRight
Cells(2, 12).HorizontalAlignment = xlRight

With Range(Cells(5, 2), Cells(5, 3))
  .Merge
  .HorizontalAlignment = xlCenter
End With
Range(Cells(6, 2), Cells(14, 2)).HorizontalAlignment = xlCenter
Range(Cells(6, 3), Cells(7, 3)).NumberFormat = "+0.0000;-0.0000"
Range(Cells(8, 3), Cells(14, 3)).NumberFormat = "+0.00000E+00;-0.00000E+00"

Range(Cells(16, 2), Cells(72, 3)).HorizontalAlignment = xlCenter
Range(Cells(17, 2), Cells(72, 2)).NumberFormat = "0.0"
Range(Cells(17, 3), Cells(72, 3)).NumberFormat = "+0.0000;-0.0000"

With Range(Cells(5, 2), Cells(72, 3))
  With .Borders(xlEdgeLeft)
    .LineStyle = xlContinuous: .Weight = xlHairline
  End With
  With .Borders(xlEdgeRight)
    .LineStyle = xlContinuous: .Weight = xlHairline
  End With
End With

ActiveSheet.Shapes.AddPicture _
  Filename:=ThisWorkbook.Path & "\Template\AspEquation.png", _
  LinkToFile:=msoFalse, SaveWithDocument:=msoTrue, _
  Top:=15.75, Left:=141, Height:=30.75, Width:=285.62

End Sub


Public Sub RenameList()

Dim snw As Long

With Application
  snw = .SheetsInNewWorkbook
  .SheetsInNewWorkbook = 1
  Workbooks.Add
  .SheetsInNewWorkbook = snw
End With
With Cells
  .ClearFormats
  .Font.Name = "ＭＳ 明朝"
  .Font.Size = 10
End With
ActiveWindow.Zoom = 100
ActiveSheet.Name = "リネームリスト"

Columns(1).ColumnWidth = 2
Columns(2).ColumnWidth = 26
Columns(3).ColumnWidth = 5
Columns(4).ColumnWidth = 58
Columns(5).ColumnWidth = 2

Cells(2, 2).Value = "■リネーム対応表"
Cells(4, 2).Value = "リネーム前"
Cells(4, 4).Value = "リネーム後"
Cells(6, 2).Value = "   ↑シート名の(ASP)は、"
Cells(7, 2).Value = "     省略可能です"
Cells(6, 3).Value = "拡張子"
Cells(7, 3).Value = "PDF"
Cells(6, 4).Value = "   ↑品目コードが<未記入>のものは参考表示。"
Cells(7, 4).Value = "     実際は未記入ではリネーム処理されません"

Cells(4, 2).HorizontalAlignment = xlCenter
Cells(4, 4).HorizontalAlignment = xlCenter
Cells(5, 3).HorizontalAlignment = xlCenter
With Range(Cells(6, 3), Cells(7, 3))
  .ShrinkToFit = True
  .HorizontalAlignment = xlCenter
End With

With Cells(7, 3).Borders
  .LineStyle = xlContinuous: .Weight = xlThin
End With

End Sub


Public Sub OldToNew()

Dim xlWb As Excel.Workbook
Dim fDialog As FileDialog
Dim objBorder As clsBorder

Dim k1 As Variant, k2 As Variant
Dim curLinks As Variant
Dim chkWb As Boolean
Dim Stt As Integer
Dim i As Long, eRow As Long
Dim curFml As String, chgFml As String
Dim oldFml(4) As String, newFml(4) As String

oldFml(1) = ",C19,C20,C21,C22,C23)"
oldFml(2) = ",C26,C27,C28,C29,C30)"
oldFml(3) = ",C19,C20,C21,C22,C23,C26,C27,C28,C29,C30)"
oldFml(4) = "PressValue()"

newFml(1) = ",C19,C20,C21,C22,C23,C24,C25)"
newFml(2) = ",C26,C27,C28,C29,C30,C31,C32)"
newFml(3) = ",C19,C20,C21,C22,C23,C24,C25,C26,C27,C28,C29,C30,C31,C32)"
newFml(4) = "PressValue(C8,C9,C13,C14,F8)"

Set fDialog = Application.FileDialog(fileDialogType:= _
                                     msoFileDialogOpen)
With fDialog
  .Title = ProjectTitle
  .ButtonName = "新版へ変換"
  .FilterIndex = 13
  .AllowMultiSelect = False
  If .Show = -1 Then
    chkWb = True
    If Workbooks.Count > 0 Then
      For Each k1 In Workbooks
        If k1.FullName = .SelectedItems(1) Then
          k1.Activate
          chkWb = False
          Exit For
        End If
      Next
    End If
    If chkWb Then
      Workbooks.Open _
        Filename:=.SelectedItems(1), UpdateLinks:=0
      DoEvents
    End If
    
    Set xlWb = ActiveWorkbook
  Else
    Set xlWb = Nothing
  End If
End With
Set fDialog = Nothing
Application.ScreenUpdating = True

If Not xlWb Is Nothing Then
  Application.ScreenUpdating = False
  
  Stt = 0
  objError.CheckSummary
  If objError.intCode = 0 Then
    curLinks = xlWb.LinkSources(xlExcelLinks)
    If Not IsEmpty(curLinks) Then
      For i = 1 To UBound(curLinks)
        If curLinks(i) <> ThisWorkbook.FullName And _
           InStr(curLinks(i), ProjectTitle) > 0 Then
          xlWb.ChangeLink _
            Name:=curLinks(i), _
            NewName:=ThisWorkbook.FullName, _
            Type:=xlLinkTypeExcelLinks
        End If
      Next
    End If
    
    Stt = 1
    Set objBorder = New clsBorder
    For Each k1 In Worksheets
      With k1
        If .Name <> DefaultSheet1 And _
           .Name <> DefaultSheet3 And _
           .Name <> "G00" And _
           Left(.Name, 1) <> "B" And _
           Left(.Name, 1) <> "M" Then
          If .Cells(2, 5).Value <> "尺度" Then
            Stt = Stt + 10
            .Cells(2, 5).Value = "尺度"
            
            With .Cells(2, 6)
              With .Validation
                .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
                     Formula1:="自動,1倍図,2倍図,5倍図,10倍図"
                .InCellDropdown = True
                .IgnoreBlank = False
                .ShowInput = False
                .ShowError = True
                .ErrorTitle = ProjectTitle
                .ErrorMessage = "リスト以外の項目は入力出来ません"
                .IMEMode = xlIMEModeNoControl
              End With
              .Value = "自動"
            End With
          End If
          
          If .Cells(26, 2).Value <> "コーニック定数2" Then
            Stt = Stt + 10
            eRow = .Cells(.Rows.Count, 2).End(xlUp).Row
            If eRow < .Cells(.Rows.Count, 3).End(xlUp).Row Then
              eRow = .Cells(.Rows.Count, 3).End(xlUp).Row
            End If
            
            .Range(.Cells(24, 2), .Cells(eRow, 3)).Cut _
              Destination:=.Range(.Cells(26, 2), .Cells(eRow + 2, 3))
            With .Range(.Cells(24, 2), .Cells(25, 3))
              .Font.Name = "ＭＳ 明朝"
              .Font.Size = 10
            End With
            .Cells(24, 2).Value = "12次の係数1"
            .Cells(25, 2).Value = "14次の係数1"
            
            If Not IsEmpty(.Cells(19, 3).Value) Then
              .Cells(24, 3).Value = 0
              .Cells(25, 3).Value = 0
            End If
            .Range(.Cells(24, 3), .Cells(25, 3)).NumberFormat = "+0.00000E+00;-0.00000E+00"
            objBorder.SetUp rngCells:=.Range(.Cells(24, 2), .Cells(25, 3))
          End If
          
          If .Cells(31, 2).Value <> "12次の係数2" Or _
             .Cells(32, 2).Value <> "14次の係数2" Then
            Stt = Stt + 10
            eRow = .Cells(.Rows.Count, 2).End(xlUp).Row
            If eRow < .Cells(.Rows.Count, 3).End(xlUp).Row Then
              eRow = .Cells(.Rows.Count, 3).End(xlUp).Row
            End If
            
            If eRow > 30 Then
              .Range(.Cells(31, 2), .Cells(eRow, 3)).Cut _
                Destination:=.Range(.Cells(33, 2), .Cells(eRow + 2, 3))
              With .Range(.Cells(31, 2), .Cells(32, 3))
                .Font.Name = "ＭＳ 明朝"
                .Font.Size = 10
              End With
            End If
            .Cells(31, 2).Value = "12次の係数2"
            .Cells(32, 2).Value = "14次の係数2"
            
            If Not IsEmpty(.Cells(26, 3).Value) Then
              .Cells(31, 3).Value = 0
              .Cells(32, 3).Value = 0
            End If
            .Range(.Cells(31, 3), .Cells(32, 3)).NumberFormat = "+0.00000E+00;-0.00000E+00"
            objBorder.SetUp rngCells:=.Range(.Cells(31, 2), .Cells(32, 3))
          End If
          
          If .Cells(26, 5).Value <> "注記6" Then
            Stt = Stt + 10
            eRow = .Cells(.Rows.Count, 5).End(xlUp).Row
            For i = 6 To 9
              If eRow < .Cells(.Rows.Count, i).End(xlUp).Row Then
                eRow = .Cells(.Rows.Count, i).End(xlUp).Row
              End If
            Next
            
            If eRow > 25 Then
              .Range(.Cells(26, 5), .Cells(eRow, 9)).Cut _
                Destination:=.Range(.Cells(29, 5), .Cells(eRow + 3, 9))
              With .Range(.Cells(26, 5), .Cells(28, 9))
                .Font.Name = "ＭＳ 明朝"
                .Font.Size = 10
              End With
            End If
            .Cells(26, 5).Value = "注記6"
            .Cells(27, 5).Value = "注記7"
            .Cells(28, 5).Value = "注記8"
            
            With .Range(.Cells(26, 6), .Cells(28, 9))
              .ShrinkToFit = True
              .Merge Across:=True
            End With
            objBorder.SetUp rngCells:=.Range(.Cells(26, 5), .Cells(28, 9))
          End If
          
          If .Cells(30, 5).Value <> "■類似硝種データ" Then
            Stt = Stt + 10
            eRow = .Cells(.Rows.Count, 5).End(xlUp).Row
            If eRow < .Cells(.Rows.Count, 6).End(xlUp).Row Then
              eRow = .Cells(.Rows.Count, 6).End(xlUp).Row
            End If
            
            If eRow > 28 Then
              .Range(.Cells(29, 5), .Cells(eRow, 6)).Cut _
                Destination:=.Range(.Cells(37, 5), .Cells(eRow + 8, 6))
              With .Range(.Cells(29, 5), .Cells(36, 6))
                .Font.Name = "ＭＳ 明朝"
                .Font.Size = 10
              End With
            End If
            .Cells(30, 5).Value = "■類似硝種データ"
            .Cells(31, 5).Value = "HOYA"
            .Cells(32, 5).Value = "OHARA"
            .Cells(33, 5).Value = "SCHOTT"
            .Cells(34, 5).Value = "光ガラス"
            .Cells(35, 5).Value = "成都光明"
            .Cells(36, 5).Value = "新華光"
            
            .Range(.Cells(31, 6), .Cells(36, 6)).HorizontalAlignment = xlRight
            objBorder.SetUp rngCells:=.Range(.Cells(31, 5), .Cells(36, 6))
          End If
          
          For Each k2 In .UsedRange.Cells
            If k2.HasFormula Then
              curFml = k2.Formula
              chgFml = k2.Formula
              If InStr(chgFml, oldFml(3)) > 0 Then
                chgFml = Replace(chgFml, oldFml(3), newFml(3))
              Else
                If InStr(chgFml, oldFml(1)) > 0 Then
                  chgFml = Replace(chgFml, oldFml(1), newFml(1))
                End If
                If InStr(chgFml, oldFml(2)) > 0 Then
                  chgFml = Replace(chgFml, oldFml(2), newFml(2))
                End If
              End If
              If InStr(chgFml, oldFml(4)) > 0 Then
                chgFml = Replace(chgFml, oldFml(4), newFml(4))
              End If
              
              If chgFml <> curFml Then
                Stt = Stt + 10
                k2.Formula = chgFml
              End If
            End If
          Next
        End If
      End With
    Next
    Set objBorder = Nothing
  End If
  
  Select Case Stt
    Case 0
      If chkWb Then xlWb.Close SaveChanges:=False
    Case 1
      objError.Message intMsgNumber:=14
      If chkWb Then xlWb.Close SaveChanges:=False
    Case Is > 1
      If chkWb Then
        objError.Message intMsgNumber:=203, _
                         strAddWord1:=xlWb.Name
        Select Case objError.intReturnValue
          Case vbOK
            xlWb.Close SaveChanges:=True
          Case vbCancel
            xlWb.Close SaveChanges:=True, Filename:=xlWb.Path & "\新版_" & xlWb.Name
        End Select
      Else
        objError.Message intMsgNumber:=104
      End If
    Case Else
      If chkWb Then xlWb.Close SaveChanges:=False
  End Select
  Set xlWb = Nothing
  
  Application.ScreenUpdating = True
End If

End Sub


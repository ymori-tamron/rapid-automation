VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsRapidText"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False


Option Explicit
Option Base 1
Option Compare Binary

Private lngIndex As Long, lngRow As Long, lngCol As Long
Private lngBasePos As Long
Private dblX As Double, dblY As Double, dblH As Double, dblW As Double
Private dblSpace As Double, dblPressWeight As Double
Private strFont As String, strItem As String


Public Sub Add()

Dim rpdTextData As zwDrawCAD.TextData
Dim objRapidPrim As clsRapidPrim

Dim f As Integer
Dim cfgFileName As String, defFont As String

objProgCtrl.AddBar lngDivCount:=2
cfgFileName = ThisWorkbook.Path & "\Config\Param.cfg"
f = FreeFile(0)
Open cfgFileName For Input As #f
Input #f, lngIndex, defFont, dblX, dblY, dblH, dblW, lngRow, lngCol
Close #f

cfgFileName = ThisWorkbook.Path & "\Config\"
If ActiveSheet.Name <> "G00" Then
  If Left(ActiveSheet.Name, 1) <> "B" And _
     Left(ActiveSheet.Name, 1) <> "M" Then
    If InStr(ActiveSheet.Name, "ASP") > 0 Then
      cfgFileName = cfgFileName & "TxtA3.cfg"
    Else
      cfgFileName = cfgFileName & "TxtA4.cfg"
    End If
  Else
    cfgFileName = cfgFileName & "TxtA4Balsam.cfg"
  End If
Else
  If rpdDoc.DrawInfo.Sheet = zwSheetA3 Then
    cfgFileName = cfgFileName & "TxtA3Assy.cfg"
  Else
    cfgFileName = cfgFileName & "TxtA4Assy.cfg"
  End If
End If
f = FreeFile(0)
Open cfgFileName For Input As #f
Do While Not EOF(f)
  Input #f, lngIndex, strItem, dblX, dblY, dblH, dblW, lngRow, lngCol
  
  Set rpdTextData = New zwDrawCAD.TextData
  Set objRapidPrim = New clsRapidPrim
  lngBasePos = 4: dblSpace = 0: strFont = defFont
  With rpdTextData
    .Text = CellToText
    .FaceName = strFont
    .Height = dblH / rpdDoc.DrawInfo.Numerator
    .Width = dblW / rpdDoc.DrawInfo.Numerator
    .Space = dblSpace
    .BasePosition = lngBasePos
    .BasePoint.x = dblX / rpdDoc.DrawInfo.Numerator
    .BasePoint.y = dblY / rpdDoc.DrawInfo.Numerator
    objRapidPrim.SetUp prmObject:=rpdTextData, lngLayerNo:=CLng(Alpha(12)), _
                       lngPenStyle:=zwStyleSolid, lngPenWidth:=zwWidthThin
    rpdDoc.DataBase.Add Prim:=rpdTextData
  End With
  Set rpdTextData = Nothing
  Set objRapidPrim = Nothing
Loop
Close #f
objProgCtrl.AddBar lngDivCount:=2

End Sub


Public Sub AddE(ByRef gpdObject As clsGlassPlotData)

Dim rpdTextData As zwDrawCAD.TextData
Dim rpdDPoint1 As zwDrawCAD.DPoint
Dim rpdDPoint2 As zwDrawCAD.DPoint
Dim objRapidPrim As clsRapidPrim

Dim chk As Boolean, esc As Boolean
Dim f As Integer
Dim i As Long, j As Long
Dim dh As Double
Dim mh(3) As Double, xx(3) As Double
Dim cfgFileName As String, defFont As String

objProgCtrl.AddBar lngDivCount:=2
cfgFileName = ThisWorkbook.Path & "\Config\Param.cfg"
f = FreeFile(0)
Open cfgFileName For Input As #f
Input #f, lngIndex, defFont, dblX, dblY, dblH, dblW, lngRow, lngCol
Close #f

cfgFileName = ThisWorkbook.Path & "\Config\"
cfgFileName = cfgFileName & "TxtA3E.cfg"
f = FreeFile(0)
Open cfgFileName For Input As #f
Do While Not EOF(f)
  Input #f, lngIndex, strItem, dblX, dblY, dblH, dblW, lngRow, lngCol
  
  If strItem = "サグデータ" Then
    Set rpdDPoint1 = New zwDrawCAD.DPoint
    Set rpdDPoint2 = New zwDrawCAD.DPoint
    If Not gpdObject.blnAsp(1) Or _
       Not gpdObject.blnAsp(2) Then
      rpdDPoint1.x = (dblX - 0.1) / rpdDoc.DrawInfo.Numerator
      rpdDPoint2.x = (dblX + 0.1) / rpdDoc.DrawInfo.Numerator
      
      rpdDPoint1.y = (dblY + 0.1) / rpdDoc.DrawInfo.Numerator
      rpdDPoint2.y = (dblY - 0.1) / rpdDoc.DrawInfo.Numerator
      rpdDoc.DataBase.SelectArea _
        FirstPoint:=rpdDPoint1, LastPoint:=rpdDPoint2, bSelect:=True
      
      rpdDoc.DataBase.RemoveSelect
      rpdDoc.Views(0).Redraw
    End If
    Set rpdDPoint1 = Nothing
    Set rpdDPoint2 = Nothing
    
    mh(1) = Fix((Cells(9, 6).Value / 2) / 0.1) * 0.1
    mh(2) = Fix((Cells(10, 6).Value / 2) / 0.1) * 0.1
    xx(1) = dblX / rpdDoc.DrawInfo.Numerator
    xx(2) = dblX / rpdDoc.DrawInfo.Numerator
    xx(3) = (dblX - 16.125) / rpdDoc.DrawInfo.Numerator
    If gpdObject.blnAsp(1) And _
       gpdObject.blnAsp(2) Then
      If mh(1) >= mh(2) Then
        mh(3) = mh(1)
      Else
        mh(3) = mh(2)
      End If
      xx(1) = (dblX - 6.375) / rpdDoc.DrawInfo.Numerator
      xx(2) = (dblX + 6.375) / rpdDoc.DrawInfo.Numerator
    Else
      If gpdObject.blnAsp(1) Then
        mh(3) = mh(1)
      Else
        mh(3) = mh(2)
      End If
    End If
    dh = Fix(((mh(3) - 0.1) / 52) / 0.1) * 0.1 + 0.1
    
    esc = False
    For j = -1 To 54
      For i = 1 To 3
        If j * dh < mh(i) + dh * 2 Then
          chk = False
          Select Case i
            Case 1, 2
              If gpdObject.blnAsp(i) Then chk = True
            Case 3
              If j > -1 Then chk = True
          End Select
          
          If chk Then
            Set rpdTextData = New zwDrawCAD.TextData
            Set objRapidPrim = New clsRapidPrim
            lngBasePos = 4: dblSpace = 0: strFont = defFont
            With rpdTextData
              If j = -1 Then
                .Text = "X/R" & CStr(i)
              Else
                Select Case i
                  Case 1
                    .Text = Format(Sag(Radius:=Cells(8, 3).Value, Diameter:=(j * dh) * 2, _
                                       Conic:=Cells(19, 3).Value, _
                                       A4:=Cells(20, 3).Value, A6:=Cells(21, 3).Value, _
                                       A8:=Cells(22, 3).Value, A10:=Cells(23, 3).Value, _
                                       A12:=Cells(24, 3).Value, A14:=Cells(25, 3).Value), "+0.0000;-0.0000")
                  Case 2
                    .Text = Format(Sag(Radius:=Cells(9, 3).Value, Diameter:=(j * dh) * 2, _
                                       Conic:=Cells(26, 3).Value, _
                                       A4:=Cells(27, 3).Value, A6:=Cells(28, 3).Value, _
                                       A8:=Cells(29, 3).Value, A10:=Cells(30, 3).Value, _
                                       A12:=Cells(31, 3).Value, A14:=Cells(32, 3).Value), "+0.0000;-0.0000")
                  Case 3
                    .Text = Format(j * dh, "0.0")
                End Select
              End If
              
              .FaceName = strFont
              .Height = dblH / rpdDoc.DrawInfo.Numerator
              .Width = dblW / rpdDoc.DrawInfo.Numerator
              .Space = dblSpace
              .BasePosition = lngBasePos
              .BasePoint.x = xx(i)
              .BasePoint.y = (dblY - 3 * (j + 1)) / rpdDoc.DrawInfo.Numerator
              objRapidPrim.SetUp prmObject:=rpdTextData, lngLayerNo:=CLng(Alpha(12)), _
                                 lngPenStyle:=zwStyleSolid, lngPenWidth:=zwWidthThin
              rpdDoc.DataBase.Add Prim:=rpdTextData
            End With
            Set rpdTextData = Nothing
            Set objRapidPrim = Nothing
          End If
        Else
          If i = 3 Then esc = True
        End If
      Next
      If esc Then Exit For
    Next
  Else
    Set rpdTextData = New zwDrawCAD.TextData
    Set objRapidPrim = New clsRapidPrim
    lngBasePos = 4: dblSpace = 0: strFont = defFont
    With rpdTextData
      .Text = CellToText
      .FaceName = strFont
      .Height = dblH / rpdDoc.DrawInfo.Numerator
      .Width = dblW / rpdDoc.DrawInfo.Numerator
      .Space = dblSpace
      .BasePosition = lngBasePos
      .BasePoint.x = dblX / rpdDoc.DrawInfo.Numerator
      .BasePoint.y = dblY / rpdDoc.DrawInfo.Numerator
      objRapidPrim.SetUp prmObject:=rpdTextData, lngLayerNo:=CLng(Alpha(12)), _
                         lngPenStyle:=zwStyleSolid, lngPenWidth:=zwWidthThin
      rpdDoc.DataBase.Add Prim:=rpdTextData
    End With
    Set rpdTextData = Nothing
    Set objRapidPrim = Nothing
  End If
Loop
Close #f
objProgCtrl.AddBar lngDivCount:=2

End Sub


Public Sub AddP(ByRef ppdObject As clsPressPlotData)

Dim rpdTextData As zwDrawCAD.TextData
Dim objRapidPrim As clsRapidPrim

Dim f As Integer
Dim j As Long
Dim d(2) As Double
Dim cfgFileName As String, defFont As String

With ppdObject
  For j = 1 To 2
    If .dblCornerH(1, j) - .dblSurfaceH(j) > 0 Then
      d(j) = .dblSurfaceH(j) * 2
    Else
      d(j) = .dblCornerH(2, j) * 2
    End If
  Next
  dblPressWeight = _
    GlassWeight(Radius1:=.dblRadius(1), Radius2:=.dblRadius(2), _
                Thickness:=.dblCenterX(2) - .dblCenterX(1), _
                SpG:=Cells(16, 9).Value, _
                Diameter1:=d(1), Diameter2:=d(2), _
                MaxDiameter:=.dblCornerH(2, 1) * 2)
End With

objProgCtrl.AddBar lngDivCount:=2
cfgFileName = ThisWorkbook.Path & "\Config\Param.cfg"
f = FreeFile(0)
Open cfgFileName For Input As #f
Input #f, lngIndex, defFont, dblX, dblY, dblH, dblW, lngRow, lngCol
Close #f

cfgFileName = ThisWorkbook.Path & "\Config\"
cfgFileName = cfgFileName & "TxtA4Press.cfg"
f = FreeFile(0)
Open cfgFileName For Input As #f
Do While Not EOF(f)
  Input #f, lngIndex, strItem, dblX, dblY, dblH, dblW, lngRow, lngCol
  
  Set rpdTextData = New zwDrawCAD.TextData
  Set objRapidPrim = New clsRapidPrim
  lngBasePos = 4: dblSpace = 0: strFont = defFont
  With rpdTextData
    .Text = CellToText
    .FaceName = strFont
    .Height = dblH / rpdDoc.DrawInfo.Numerator
    .Width = dblW / rpdDoc.DrawInfo.Numerator
    .Space = dblSpace
    .BasePosition = lngBasePos
    .BasePoint.x = dblX / rpdDoc.DrawInfo.Numerator
    .BasePoint.y = dblY / rpdDoc.DrawInfo.Numerator
    objRapidPrim.SetUp prmObject:=rpdTextData, lngLayerNo:=lngPre(3), _
                       lngPenStyle:=zwStyleSolid, lngPenWidth:=zwWidthThin
    rpdDoc.DataBase.Add Prim:=rpdTextData
  End With
  Set rpdTextData = Nothing
  Set objRapidPrim = Nothing
Loop
Close #f
objProgCtrl.AddBar lngDivCount:=2

End Sub


Private Function CellToText() As String

Dim maxChr As Long
Dim i As Long, cc As Long, maxCol As Long
Dim txt As String

If lngRow = 0 And _
   lngCol = 0 Then
  Select Case strItem
    Case "尺度"
      CellToText = rpdDoc.DrawInfo.Numerator
    Case "コート1M"
      dblX = dblX - dblW * 1.25 - _
        (Cells(10, 3).Value / 2) * rpdDoc.DrawInfo.Numerator
      dblY = dblY + dblH * 1.25
      If IsEmpty(Cells(10, 9).Value) Or _
         Cells(10, 9).Value = "-" Then
        CellToText = " "
      Else
        CellToText = "※"
      End If
    Case "コート2M"
      dblX = dblX + dblW * 1.25 + _
        (Cells(10, 3).Value / 2) * rpdDoc.DrawInfo.Numerator
      dblY = dblY + dblH * 1.25
      If IsEmpty(Cells(11, 9).Value) Or _
         Cells(11, 9).Value = "-" Then
        CellToText = " "
      Else
        CellToText = "※"
      End If
    Case "重量Z"
      CellToText = Format(dblPressWeight, "0.000""g""")
    Case "改定番号"
      If Not CBool(fCheck(5)) Then
        CellToText = "00"
      Else
        CellToText = " "
      End If
    Case Else
      CellToText = " "
  End Select
ElseIf lngRow > 0 And _
       lngCol > 0 Then
  If Not IsEmpty(Cells(lngRow, lngCol).Value) Then
    Select Case strItem
      Case "作成日Y"
        CellToText = Year(Cells(lngRow, lngCol).Value)
      Case "作成日M"
        CellToText = Month(Cells(lngRow, lngCol).Value)
      Case "作成日D"
        CellToText = Day(Cells(lngRow, lngCol).Value)
      Case "データ印設計Y"
        strFont = "ＭＳ 明朝"
        CellToText = Year(Cells(lngRow, lngCol).Value) & "." & Space(8)
      Case "データ印設計M"
        strFont = "ＭＳ 明朝"
        If Month(Cells(lngRow, lngCol).Value) < 10 Then
          dblW = dblW * (4 / 3)
        End If
        CellToText = Space(3) & Month(Cells(lngRow, lngCol).Value) & "." & Space(2)
      Case "データ印設計D"
        strFont = "ＭＳ 明朝"
        If Day(Cells(lngRow, lngCol).Value) < 10 Then
          CellToText = Space(6) & "-" & Day(Cells(lngRow, lngCol).Value)
        Else
          CellToText = Space(6) & Day(Cells(lngRow, lngCol).Value)
        End If
      Case "データ印製図Y"
        strFont = "ＭＳ 明朝"
        CellToText = Year(Cells(lngRow, lngCol).Value) & "." & Space(8)
      Case "データ印製図M"
        strFont = "ＭＳ 明朝"
        If Month(Cells(lngRow, lngCol).Value) < 10 Then
          dblW = dblW * (4 / 3)
        End If
        CellToText = Space(3) & Month(Cells(lngRow, lngCol).Value) & "." & Space(2)
      Case "データ印製図D"
        strFont = "ＭＳ 明朝"
        If Day(Cells(lngRow, lngCol).Value) < 10 Then
          CellToText = Space(6) & "-" & Day(Cells(lngRow, lngCol).Value)
        Else
          CellToText = Space(6) & Day(Cells(lngRow, lngCol).Value)
        End If
      Case "メーカー名"
        If Cells(lngRow, lngCol).Value <> "OTHER" Then
          CellToText = "(" & Cells(lngRow, lngCol).Value & ")"
        Else
          CellToText = " "
        End If
      Case "コート有効径1", "コート有効径2"
        If Not IsNumeric(Cells(lngRow, lngCol).Value) Then
          dblH = dblH * 0.875
          dblW = dblW * 0.875
        End If
        CellToText = Cells(lngRow, lngCol).Text
      Case "コーニック定数1", "コーニック定数2"
        CellToText = Format(Cells(lngRow, lngCol).Value + 1, "+0.0000;-0.0000")
      Case "品目コード"
        dblSpace = (dblW / 4) / rpdDoc.DrawInfo.Numerator
        If CBool(fCheck(5)) Then
          CellToText = Left(Trim(Cells(lngRow, lngCol).Value), 11) & Space(3)
        Else
          CellToText = Left(Trim(Cells(lngRow, lngCol).Value), 13) & Space(1)
        End If
      Case "品目コードA0"
        dblSpace = (dblW / 4) / rpdDoc.DrawInfo.Numerator
        If InStr(Cells(lngRow, lngCol).Value, ",") > 0 Then
          txt = Trim(Left(Cells(lngRow, lngCol).Value, _
                    InStr(Cells(lngRow, lngCol).Value, ",") - 1))
        Else
          txt = Trim(Cells(lngRow, lngCol).Value)
        End If
        If CBool(fCheck(5)) Then
          CellToText = Left(txt, 11) & Space(3)
        Else
          CellToText = Left(txt, 13) & Space(1)
        End If
      Case "品目コードA1"
        dblSpace = (dblW / 4) / rpdDoc.DrawInfo.Numerator
        If InStr(Cells(lngRow, lngCol).Value, ",") > 0 Then
          txt = Trim(Mid(Cells(lngRow, lngCol).Value, _
                   InStr(Cells(lngRow, lngCol).Value, ",") + 1))
        Else
          txt = Trim(Cells(lngRow, lngCol).Value)
        End If
        If CBool(fCheck(5)) Then
          CellToText = Left(txt, 11) & Space(3)
        Else
          CellToText = Left(txt, 13) & Space(1)
        End If
      Case "品名(上段)"
        lngBasePos = 1
        If Cells(lngRow, lngCol).Value = "(機種の特徴を記入)" Then
          CellToText = " "
        Else
          txt = Cells(lngRow, lngCol).Text
          If LenB(StrConv(txt, vbFromUnicode)) > 20 Then
            dblW = dblW * (20 / LenB(StrConv(txt, vbFromUnicode)))
          End If
          CellToText = txt
        End If
      Case "品名(下段)"
        lngBasePos = 7
        If InStr(Cells(lngRow, lngCol).Value, ",") > 0 Then
          txt = Trim(Left(Cells(lngRow, lngCol).Value, _
                    InStr(Cells(lngRow, lngCol).Value, ",") - 1))
          If LenB(StrConv(txt, vbFromUnicode)) > 20 Then
            dblW = dblW * (20 / LenB(StrConv(txt, vbFromUnicode)))
          End If
          CellToText = txt
        Else
          CellToText = " "
        End If
      Case "部品番号A0"
        If InStr(Cells(lngRow, lngCol).Value, ",") > 0 Then
          txt = Trim(Left(Cells(lngRow, lngCol).Value, _
                    InStr(Cells(lngRow, lngCol).Value, ",") - 1))
        Else
          txt = Trim(Cells(lngRow, lngCol).Value)
        End If
        CellToText = txt
      Case "部品番号A1"
        lngBasePos = 1
        If InStr(Cells(lngRow, lngCol).Value, ",") > 0 Then
          txt = Trim(Mid(Cells(lngRow, lngCol).Value, _
                   InStr(Cells(lngRow, lngCol).Value, ",") + 1))
        Else
          txt = Trim(Cells(lngRow, lngCol).Value)
        End If
        CellToText = txt
      Case "部品番号A2"
        lngBasePos = 7
        If InStr(Cells(lngRow, lngCol).Value, ",") > 0 Then
          txt = Trim(Mid(Cells(lngRow, lngCol).Value, _
                   InStr(Cells(lngRow, lngCol).Value, ",") + 1))
          If LenB(StrConv(txt, vbFromUnicode)) > 12 Then
            dblW = dblW * (12 / LenB(StrConv(txt, vbFromUnicode)))
          End If
          CellToText = txt
        Else
          CellToText = " "
        End If
      Case "注記1", "注記2", "注記3", "注記4", "注記5", _
           "注記6", "注記7", "注記8", "注記9", "注記10"
        If rpdDoc.DrawInfo.Sheet = zwSheetA3 Then
          maxChr = 86
        Else
          maxChr = 59
        End If
        
        If LenB(StrConv(Cells(lngRow, lngCol).Text, vbFromUnicode)) > maxChr Then
          dblW = dblW * (maxChr / LenB(StrConv(Cells(lngRow, lngCol).Text, vbFromUnicode)))
        End If
        lngBasePos = 3
        CellToText = Cells(lngRow, lngCol).Text
      Case "注記1A", "注記2A", "注記3A", "注記4A", "注記5A", _
           "注記6A", "注記7A", "注記8A", "注記9A", "注記10A"
        If LenB(StrConv(Cells(lngRow, lngCol).Text, vbFromUnicode)) > 60 Then
          dblW = dblW * (60 / LenB(StrConv(Cells(lngRow, lngCol).Text, vbFromUnicode)))
        End If
        lngBasePos = 3
        CellToText = Cells(lngRow, lngCol).Text
      Case "部品番号B1"
        If InStr(Cells(lngRow, lngCol).Value, "接合図") > 0 Then
          lngBasePos = 1
          CellToText = Left(Trim(Cells(lngRow, lngCol).Value), 3)
        Else
          CellToText = Cells(lngRow, lngCol).Text
        End If
      Case "部品番号B2"
        If InStr(Cells(lngRow, lngCol).Value, "接合図") > 0 Then
          lngBasePos = 7
          txt = Mid(Trim(Cells(lngRow, lngCol).Value), 4)
          If LenB(StrConv(txt, vbFromUnicode)) > 12 Then
            dblW = dblW * (12 / LenB(StrConv(txt, vbFromUnicode)))
          End If
          CellToText = txt
        Else
          CellToText = " "
        End If
      Case "部品番号Z"
        CellToText = "Z" & Mid(Trim(Cells(lngRow, lngCol).Value), 2, 2)
      Case "品目コードZ"
        dblSpace = (dblW / 4) / rpdDoc.DrawInfo.Numerator
        If Len(Trim(Cells(lngRow, lngCol).Value)) > 14 Then
          If CBool(fCheck(5)) Then
            CellToText = Left(Right(Trim(Cells(lngRow, lngCol).Value), 13), 11) & Space(3)
          Else
            CellToText = Right(Trim(Cells(lngRow, lngCol).Value), 13) & Space(1)
          End If
        Else
          CellToText = " "
        End If
      Case Else
        CellToText = Cells(lngRow, lngCol).Text
    End Select
  Else
    CellToText = " "
  End If
ElseIf lngRow > 0 And _
       lngCol = -1 Then
  With Worksheets(DefaultSheet3)
    cc = 0
    maxCol = .Cells(6, .Columns.Count).End(xlToLeft).Column
    If maxCol > 2 Then
      For i = 3 To maxCol Step 3
        If .Cells(6, i).Value = ActiveSheet.Name Then
          cc = i: Exit For
        End If
      Next
    End If
    If cc > 0 Then
      If Not IsEmpty(.Cells(lngRow, cc).Value) Then
        Select Case strItem
          Case "ニュートン1S", "ニュートン2S"
            If Not IsEmpty(.Cells(lngRow, cc + 1).Value) Then
              CellToText = " "
            Else
              If InStr(ActiveSheet.Name, "ASP") > 0 Then
                CellToText = .Cells(lngRow, cc).Text
              Else
                CellToText = "-" & .Cells(lngRow, cc).Value & .Cells(lngRow, cc).Text
              End If
            End If
          Case "アス1S", "アス2S"
            If Not IsEmpty(.Cells(lngRow, cc + 1).Value) Then
              CellToText = " "
            Else
              CellToText = "(" & .Cells(lngRow, cc).Text & ")"
            End If
          Case "テキスト1", "テキスト2"
            If Not IsEmpty(.Cells(lngRow, cc + 1).Value) Then
              CellToText = "別途指示"
            Else
              CellToText = " "
            End If
          Case "ニュートン1A", "ニュートン2A", "アス1A", "アス2A"
            If Not IsEmpty(.Cells(lngRow, cc + 1).Value) Then
              CellToText = .Cells(lngRow, cc).Text & "相当"
            Else
              CellToText = " "
            End If
          Case "ニュートン1μ", "ニュートン2μ", "アス1μ", "アス2μ"
            If Not IsEmpty(.Cells(lngRow, cc + 1).Value) Then
              CellToText = "±" & .Cells(lngRow, cc + 1).Text & "m"
            Else
              CellToText = " "
            End If
          Case "ＮdT"
            CellToText = Format(.Cells(lngRow, cc + 1).Value, "General Number")
          Case "νdT"
            CellToText = Format(.Cells(lngRow, cc + 1).Value, "0.0#")
          Case "外径基準1", "外径基準2"
            If Left(.Cells(lngRow, cc).Value, 1) = "G" Then
              CellToText = Left(.Cells(lngRow, cc).Value, 3)
            Else
              CellToText = " "
            End If
          Case "基準面1", "基準面2"
            CellToText = "R" & CStr(.Cells(lngRow, cc).Value * (-1) + 3)
          Case "型保証偏芯S"
            CellToText = .Cells(lngRow, cc).Text & "m"
          Case "反射芯M", "型保証偏芯T"
            CellToText = .Cells(lngRow, cc + 1).Text
          Case Else
            CellToText = .Cells(lngRow, cc).Text
        End Select
      Else
        CellToText = " "
      End If
    Else
      CellToText = " "
    End If
  End With
Else
  CellToText = " "
End If

End Function


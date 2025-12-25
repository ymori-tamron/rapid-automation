VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsRapidLayout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False


Option Explicit
Option Base 1
Option Compare Binary


Public Sub AddA(ByRef apdObject As clsAssyPlotData)

Dim rpdPolyData As zwDrawCAD.PolyLineData
Dim rpdDimeData As zwDrawCAD.DimensionData
Dim rpdPickData As zwDrawCAD.PickupData
Dim rpdTextData As zwDrawCAD.TextData
Dim rpdDPoint1 As zwDrawCAD.DPoint
Dim rpdDPoint2 As zwDrawCAD.DPoint
Dim rpdDPoint3 As zwDrawCAD.DPoint
Dim rpdDPoint4 As zwDrawCAD.DPoint
Dim objRapidPrim As clsRapidPrim
Dim objRapidParam As clsRapidParam
Dim objRapidTol As clsRapidTol

Dim unit As Boolean
Dim i As Long, j As Long, rc As Long
Dim dd As Long, bc As Long
Dim ic As Long, n1 As Long, n2 As Long
Dim pn As Long, pc As Long
Dim maxItem As Long
Dim maxRow(3) As Long, ff(2) As Long
Dim aa As Double, xx As Double, yy As Double
Dim hh As Double, ww As Double, tt As Double
Dim tmp As Double, stop1 As Double, stop2 As Double
Dim area1 As Double, area2 As Double
Dim zz1(99, 2) As Double, zz2(99, 2) As Double
Dim sName As String, eName As String, txt As String
Dim strAName1(99, 2) As String
Dim strAName2(99, 2) As String

Set objRapidParam = New clsRapidParam
objRapidParam.Import strConfig:="Param"
With objRapidParam
  aa = .dblXY(1, 1) / rpdDoc.DrawInfo.Numerator
  
  rc = 0: tmp = 0
  For i = 1 To .lngMaxCount
    If .strItem(i) = "åQä‘äuìô" Then
      yy = .dblXY(i, 2) / rpdDoc.DrawInfo.Numerator
      hh = .dblHW(i, 1) / rpdDoc.DrawInfo.Numerator
      ww = .dblHW(i, 2) / rpdDoc.DrawInfo.Numerator
      dd = .lngVal(i, 1)
      bc = .lngVal(i, 2)
      rc = i: Exit For
    End If
    
    If .strItem(i) = "è≈ì_ãóó£ìô" Then
      tmp = (.dblXY(i, 2) - .dblHW(i, 1) * 2.2) / _
            rpdDoc.DrawInfo.Numerator
    End If
  Next
End With

If rc = 0 Then
  objProgCtrl.AddBar lngDivCount:=1
  Set objRapidParam = Nothing
  Exit Sub
End If

With apdObject
  If .lngMaxCount > 0 Then
    maxItem = 0
    maxRow(1) = Cells(7, 2).End(xlDown).Row
    If maxRow(1) > 7 Then
      For i = 8 To maxRow(1)
        ic = i - 7
        strAName1(ic, 1) = ""
        If ic > 1 Then strAName1(ic - 1, 2) = ""
        For j = 1 To .lngMaxCount
          If ic = 1 Then
            If Left(.strEName(j), 3) = Right(Cells(i, 2).Value, 3) Then
              strAName1(ic, 1) = .strEName(j)
              Exit For
            End If
          Else
            If Left(.strEName(j), 3) = Left(Cells(i, 2).Value, 3) Then
              strAName1(ic - 1, 2) = .strEName(j)
              If strAName1(ic, 1) <> "" Then Exit For
            End If
            If Left(.strEName(j), 3) = Right(Cells(i, 2).Value, 3) Then
              strAName1(ic, 1) = .strEName(j)
              If strAName1(ic - 1, 2) <> "" Then Exit For
            End If
          End If
        Next
        
        zz1(ic, 1) = .dblCenterX - .dblMaxLength / 2 + _
                     .varElement(strAName1(ic, 1), "POS")
        If ic > 1 Then
          zz1(ic - 1, 2) = .dblCenterX - .dblMaxLength / 2 + _
                           .varElement(strAName1(ic - 1, 2), "POS")
        End If
      Next
      maxItem = ic
      strAName1(ic, 2) = "IMG"
      zz1(ic, 2) = .dblCenterX + .dblMaxLength / 2 + _
                   20 / rpdDoc.DrawInfo.Numerator
    End If
    
    If Not IsEmpty(Cells(maxRow(1) + 3, 2).Value) Then
      maxRow(2) = Cells(maxRow(1) + 2, 2).End(xlDown).Row
      If maxRow(2) > maxRow(1) + 2 Then
        For i = maxRow(1) + 3 To maxRow(2)
          ic = i - (maxRow(1) + 2)
          strAName2(ic, 1) = ""
          strAName2(ic, 2) = ""
          For j = 1 To .lngMaxCount
            If Left(.strEName(j), 3) = Left(Cells(i, 2).Value, 3) Then
              strAName2(ic, 1) = .strEName(j)
              If strAName2(ic, 2) <> "" Then Exit For
            End If
            If Left(.strEName(j), 3) = Right(Cells(i, 2).Value, 3) Then
              strAName2(ic, 2) = .strEName(j)
              If strAName2(ic, 1) <> "" Then Exit For
            End If
          Next
          
          zz2(ic, 1) = .dblCenterX - .dblMaxLength / 2 + _
                       .varElement(strAName2(ic, 1), "POS")
          zz2(ic, 2) = .dblCenterX - .dblMaxLength / 2 + _
                       .varElement(strAName2(ic, 2), "POS")
        Next
        maxItem = maxItem + ic
        zz2(ic + 1, 1) = 10000000000#
        zz2(ic + 1, 2) = 10000000000#
      End If
    End If
    
    If maxItem = 0 Then
      objProgCtrl.AddBar lngDivCount:=1
      Set objRapidParam = Nothing
      Exit Sub
    End If
  Else
    objProgCtrl.AddBar lngDivCount:=1
    Set objRapidParam = Nothing
    Exit Sub
  End If
End With

stop1 = 0: stop2 = 0
maxRow(3) = Cells(23, 5).End(xlDown).Row
If maxRow(3) > 23 Then
  For i = 24 To maxRow(3)
    If Cells(i, 5).Value = "çiÇËåa" Then
      If Not IsEmpty(Cells(i, 6).Value) And _
         IsNumeric(Cells(i, 6).Value) Then
        stop1 = Cells(i, 6).Value
        For j = 7 To 8
          If Not IsEmpty(Cells(i, j).Value) Then
            If Cells(i, j).Value <> stop1 Then
              stop2 = Cells(i, j).Value
            End If
          Else
            Exit For
          End If
        Next
      End If
      Exit For
    End If
  Next
End If

Set rpdDPoint1 = New zwDrawCAD.DPoint
Set rpdDPoint2 = New zwDrawCAD.DPoint
Set rpdDPoint3 = New zwDrawCAD.DPoint
Set rpdDPoint4 = New zwDrawCAD.DPoint
Set objRapidPrim = New clsRapidPrim
With apdObject
  If tmp = 0 Then
    pn = 2
  Else
    pn = Fix((tmp - .dblCenterY - .dblMaxDiameter / 2 - _
             hh * 0.8) / (hh * 1.2)) - 1
  End If
  
  n1 = 1: n2 = 1
  pc = pn: tmp = 0
  For i = 1 To maxItem
    objProgCtrl.AddBar lngDivCount:=maxItem
    If zz1(n1, 1) > zz2(n2, 1) Then
      sName = strAName2(n2, 1)
      eName = strAName2(n2, 2)
      rpdDPoint1.x = zz2(n2, 1)
      rpdDPoint2.x = zz2(n2, 2)
      unit = False
      n2 = n2 + 1
    Else
      sName = strAName1(n1, 1)
      eName = strAName1(n1, 2)
      rpdDPoint1.x = zz1(n1, 1)
      rpdDPoint2.x = zz1(n1, 2)
      unit = True
      n1 = n1 + 1
    End If
    
    If sName = "STO" Or _
       Left(sName, 2) = "AP" Then
      rpdDPoint1.y = .dblCenterY + .varElement(sName, "DIA") / 2 + aa * 2
    Else
      rpdDPoint1.x = rpdDPoint1.x + Worksheets(sName).Cells(10, 3).Value
      rpdDPoint1.y = .dblCenterY
    End If
    If eName = "STO" Or _
       Left(eName, 2) = "AP" Then
      rpdDPoint2.y = .dblCenterY + .varElement(eName, "DIA") / 2 + aa * 2
    Else
      rpdDPoint2.y = .dblCenterY
    End If
    
    If unit Then
      If eName = "IMG" Then
        txt = " BF "
      Else
        txt = " L" & Mid(CStr(n1 - 1) & Space(2), 1, 2)
      End If
      xx = Len(txt) * (ww / 4) + aa + ww / 8
      ic = 0
    Else
      txt = Format(rpdDPoint2.x - rpdDPoint1.x, "0." & String(dd, "0"))
      xx = Len(txt) * (ww / 4) + aa + ww / 8
      tt = 0
      ic = (n2 - 1) + (maxRow(1) + 2)
      If InStr(Cells(ic, 3).Value, "/") > 0 Then
        If IsNumeric(Left(Cells(ic, 3).Value, InStr(Cells(ic, 3).Value, "/") - 1)) And _
           IsNumeric(Mid(Cells(ic, 3).Value, InStr(Cells(ic, 3).Value, "/") + 1)) Then
          xx = xx + bc * (ww / 4)
          tt = (Len(txt) - bc) * (ww / 4)
          txt = txt & Space(bc)
        End If
      End If
    End If
    
    objRapidParam.SetUp strParam:="DimeParam"
    rpdDoc.DimensionParam.Height = objRapidParam.dblHW(rc, 1)
    rpdDoc.DimensionParam.Width = objRapidParam.dblHW(rc, 2)
    rpdDoc.DimensionParam.DecimalDigit = 0
    If rpdDPoint1.x < tmp Then
      pc = pc - 1
      If pc < 0 Then pc = pn
    Else
      pc = pn
    End If
    If (rpdDPoint2.x - rpdDPoint1.x) / 2 < xx Then
      If rpdDPoint1.x - tmp < aa * 2.5 Then
        rpdDoc.DimensionParam.StartMarkType = zwMarkBlackCircle
        rpdDoc.DimensionParam.StartDistance = 0.001 / rpdDoc.DrawInfo.Numerator
      End If
      rpdDPoint3.x = rpdDPoint2.x + xx
      rpdDPoint3.y = .dblCenterY + .dblMaxDiameter / 2 + _
                     hh * 0.4 + (hh * 1.2) * pc
      tmp = rpdDPoint2.x + xx * 2
    Else
      rpdDPoint3.x = (rpdDPoint1.x + rpdDPoint2.x) / 2
      If unit Then
        rpdDPoint1.y = .dblCenterY
        rpdDPoint2.y = .dblCenterY
        If eName = "IMG" Then
          rpdDPoint3.y = .dblCenterY + aa * 5 + (hh / 2) * 1.1
        Else
          rpdDPoint3.y = .dblCenterY
        End If
      Else
        rpdDPoint3.y = .dblCenterY + .dblMaxDiameter / 2 + _
                       hh * 0.4 + (hh * 1.2) * pc
        tmp = rpdDPoint2.x
      End If
    End If
    
    Set rpdDimeData = New zwDrawCAD.DimensionData
    Set objRapidTol = New clsRapidTol
    rpdDimeData.CreateSingle doc:=rpdDoc, _
                             PickStart:=rpdDPoint1, PickEnd:=rpdDPoint2, _
                             PassPoint:=rpdDPoint3, dAngle:=0
    
    rpdDimeData.Text(0) = txt
    If unit Then
      rpdDPoint3.y = rpdDPoint3.y + (hh / 2) * 1.1
    End If
    rpdDPoint3.y = rpdDPoint3.y + 0.001 / rpdDoc.DrawInfo.Numerator
    rpdDimeData.SetTextPoint doc:=rpdDoc, lIndex:=0, _
                             PutPoint:=rpdDPoint3, Arrange:=zwDimTextFree
    
    objRapidPrim.SetUp prmObject:=rpdDimeData, lngLayerNo:=CLng(Alpha(13)), _
                       lngPenStyle:=zwStyleSolid, lngPenWidth:=zwWidthThin
    rpdDoc.DataBase.Add Prim:=rpdDimeData
    objRapidTol.AddA rpObject:=objRapidParam, _
                     lngPRow:=rc + 1, lngTRow:=ic, _
                     dblBaseX:=rpdDPoint3.x, dblBaseY:=rpdDPoint3.y, _
                     dblShift:=tt, dblAngle:=0
    Set rpdDimeData = Nothing
    Set objRapidTol = Nothing
    
    If unit Then
      rpdDPoint3.y = rpdDPoint3.y - (hh / 2) * 1.1
      
      Set rpdPolyData = New zwDrawCAD.PolyLineData
      rpdPolyData.CreateMark Type:=zwMarkCross, _
                             PutPoint:=rpdDPoint3, dSize:=hh
      objRapidPrim.SetUp prmObject:=rpdPolyData, lngLayerNo:=CLng(Alpha(13)), _
                         lngPenStyle:=zwStyleSolid, lngPenWidth:=zwWidthThin
      rpdDoc.DataBase.Add Prim:=rpdPolyData
      Set rpdPolyData = Nothing
      
      Set rpdPolyData = New zwDrawCAD.PolyLineData
      rpdPolyData.CreateMark Type:=zwMarkCircle, _
                             PutPoint:=rpdDPoint3, dSize:=hh / 2
      objRapidPrim.SetUp prmObject:=rpdPolyData, lngLayerNo:=CLng(Alpha(13)), _
                         lngPenStyle:=zwStyleSolid, lngPenWidth:=zwWidthThin
      rpdDoc.DataBase.Add Prim:=rpdPolyData
      Set rpdPolyData = Nothing
    End If
  Next
  
  
  ff(1) = 0
  For i = 1 To .lngMaxCount
    If .strEName(i) = "STO" Or _
       Left(.strEName(i), 2) = "AP" Then
      If ff(1) > 0 Then
        ff(1) = -1
      Else
        ff(1) = 1
      End If
      
      ff(2) = 0
      If i > 1 Then
        If .strEName(i - 1) = "STO" Or _
           Left(.strEName(i - 1), 2) = "AP" Then
          area1 = .varElement(.strEName(i), "POS") - _
                  .varElement(.strEName(i - 1), "POS")
        Else
          area1 = .varElement(.strEName(i), "POS") - _
                  .varElement(.strEName(i - 1), "POS") - _
                  Worksheets(.strEName(i - 1)).Cells(10, 3).Value
        End If
      Else
        ff(2) = -1
      End If
      If i < .lngMaxCount Then
        area2 = .varElement(.strEName(i + 1), "POS") - _
                .varElement(.strEName(i), "POS")
      Else
        ff(2) = 1
      End If
      If ff(2) = 0 Then
        If area1 > area2 Then
          ff(2) = -1
        Else
          ff(2) = 1
        End If
      End If
      
      txt = Format(.varElement(.strEName(i), "DIA"), "0.0##")
      If .strEName(i) = "STO" Then
        If stop2 = 0 Then
          xx = (Len(txt) + 8) * (ww / 4)
        Else
          txt = txt & "Å`" & Format(stop2, "0.0##")
          xx = (Len(txt) + 9) * (ww / 4)
        End If
        txt = "çiÇËåaÉ”" & txt
      Else
        xx = (Len(txt) + 11) * (ww / 4)
        txt = "åıê¸∂ØƒåaÉ”" & txt
      End If
      
      rpdDPoint1.x = .dblCenterX - .dblMaxLength / 2 + _
                     .varElement(.strEName(i), "POS")
      rpdDPoint1.y = .dblCenterY + _
                     (.varElement(.strEName(i), "DIA") / 2 + aa * 2) * ff(1)
      rpdDPoint2.x = rpdDPoint1.x + (ww * 1.5) * ff(2)
      rpdDPoint2.y = rpdDPoint1.y + yy * ff(1)
      rpdDPoint3.x = rpdDPoint1.x + (ww * 1.5 + xx * 2) * ff(2)
      rpdDPoint3.y = rpdDPoint1.y + yy * ff(1)
      rpdDPoint4.x = rpdDPoint1.x + (ww * 1.5 + xx) * ff(2)
      rpdDPoint4.y = rpdDPoint1.y + yy * ff(1)
      
      objRapidParam.SetUp strParam:="PickParam"
      rpdDoc.PickupParam.Height = objRapidParam.dblHW(rc, 1)
      rpdDoc.PickupParam.Width = objRapidParam.dblHW(rc, 2)
      rpdDoc.PickupParam.MarkType = zwMarkNone
      
      Set rpdPickData = New zwDrawCAD.PickupData
      rpdPickData.CreateEx doc:=rpdDoc, _
                           StartPoint:=rpdDPoint1, CornerPoint:=rpdDPoint2, _
                           EndPoint:=rpdDPoint3, TextPoint:=rpdDPoint4, _
                           sText:=txt, Arrange:=zwPickupTextFreeCorrect
      objRapidPrim.SetUp prmObject:=rpdPickData, lngLayerNo:=CLng(Alpha(13)), _
                         lngPenStyle:=zwStyleSolid, lngPenWidth:=zwWidthThin
      rpdDoc.DataBase.Add Prim:=rpdPickData
      Set rpdPickData = Nothing
      
      Set rpdTextData = New zwDrawCAD.TextData
      rpdTextData.Text = "Åô"
      rpdTextData.FaceName = objRapidParam.strItem(1)
      rpdTextData.Height = hh
      rpdTextData.Width = ww
      rpdTextData.BasePosition = ff(2) * (-1) + 1
      rpdTextData.BasePoint.x = rpdDPoint3.x - (ww * 0.75) * ff(2)
      rpdTextData.BasePoint.y = rpdDPoint3.y + hh
      objRapidPrim.SetUp prmObject:=rpdTextData, lngLayerNo:=CLng(Alpha(12)), _
                         lngPenStyle:=zwStyleSolid, lngPenWidth:=zwWidthThin
      rpdDoc.DataBase.Add Prim:=rpdTextData
      Set rpdTextData = Nothing
    End If
  Next
  
  objRapidParam.SetUp strParam:="DimeParam"
  objRapidParam.SetUp strParam:="PickParam"
End With
Set rpdDPoint1 = Nothing
Set rpdDPoint2 = Nothing
Set rpdDPoint3 = Nothing
Set rpdDPoint4 = Nothing
Set objRapidPrim = Nothing
Set objRapidParam = Nothing

End Sub


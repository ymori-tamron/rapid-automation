VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsRapidFooter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False


Option Explicit
Option Base 1
Option Compare Binary


Public Sub AddA(ByRef apdObject As clsAssyPlotData)

Dim rpdLineData As zwDrawCAD.LineData
Dim rpdCircleData As zwDrawCAD.CircleData
Dim rpdArcData As zwDrawCAD.ArcData
Dim rpdBalloonData As zwDrawCAD.BalloonData
Dim rpdTextData As zwDrawCAD.TextData
Dim rpdDPoint1 As zwDrawCAD.DPoint
Dim rpdDPoint2 As zwDrawCAD.DPoint
Dim objRapidPrim As clsRapidPrim
Dim objRapidParam As clsRapidParam

Dim i As Long, n As Long, rc As Long
Dim ff(4, 6) As Long
Dim yy As Double, hh As Double, ww As Double
Dim xx(99) As Double
Dim tmpx As Double, tmpy As Double
Dim strFont As String, strGName(99) As String

ff(1, 1) = 1: ff(1, 2) = 1: ff(1, 3) = 0: ff(1, 4) = 1: ff(1, 5) = 1: ff(1, 6) = 0
ff(2, 1) = -1: ff(2, 2) = 1: ff(2, 3) = -1: ff(2, 4) = 0: ff(2, 5) = 0: ff(2, 6) = 1
ff(3, 1) = -1: ff(3, 2) = -1: ff(3, 3) = 0: ff(3, 4) = -1: ff(3, 5) = -1: ff(3, 6) = 0
ff(4, 1) = 1: ff(4, 2) = -1: ff(4, 3) = 1: ff(4, 4) = 0: ff(4, 5) = 0: ff(4, 6) = -1

Set objRapidParam = New clsRapidParam
objRapidParam.Import strConfig:="Param"
With objRapidParam
  strFont = .strItem(1)
  
  rc = 0
  For i = 1 To .lngMaxCount
    If .strItem(i) = "部品番号" Then
      yy = .dblXY(i, 2) / rpdDoc.DrawInfo.Numerator
      hh = .dblHW(i, 1) / rpdDoc.DrawInfo.Numerator
      ww = .dblHW(i, 2) / rpdDoc.DrawInfo.Numerator
      
      .SetUp strParam:="PickParam"
      rpdDoc.PickupParam.Height = .dblHW(i, 1)
      rpdDoc.PickupParam.Width = .dblHW(i, 2)
      rpdDoc.PickupParam.MarkType = zwMarkBlackCircle
      rpdDoc.PickupParam.BalloonType = zwBalloonCircle
      
      rc = i: Exit For
    End If
  Next
End With
Set objRapidParam = Nothing

If rc = 0 Then
  objProgCtrl.AddBar lngDivCount:=1
  Exit Sub
End If

With apdObject
  If .lngMaxCount > 0 Then
    lngGla = 0
    For i = 1 To .lngMaxCount
      If Left(.strEName(i), 1) = "G" Then
        lngGla = lngGla + 1
        strGName(lngGla) = .strEName(i)
        xx(lngGla) = .dblCenterX - .dblMaxLength / 2 + _
                     .varElement(.strEName(i), "POS") + _
                     .varElement(.strEName(i), "PPX")
      End If
    Next
    
    If lngGla = 0 Then
      objProgCtrl.AddBar lngDivCount:=1
      Exit Sub
    End If
  Else
    objProgCtrl.AddBar lngDivCount:=1
    Exit Sub
  End If
End With

Set rpdLineData = New zwDrawCAD.LineData
Set rpdDPoint1 = New zwDrawCAD.DPoint
Set rpdDPoint2 = New zwDrawCAD.DPoint
Set objRapidPrim = New clsRapidPrim
With apdObject
  If lngGla > 1 Then
    For i = Fix(lngGla / 2) To 1 Step -1
      If xx(i + 1) - xx(i) < (hh * (Sqr(2) + 0.1)) * 2.2 Then
        xx(i) = xx(i + 1) - (hh * (Sqr(2) + 0.1)) * 2.2
      End If
    Next
    For i = Fix(lngGla / 2) + 1 To lngGla
      If xx(i) - xx(i - 1) < (hh * (Sqr(2) + 0.1)) * 2.2 Then
        xx(i) = xx(i - 1) + (hh * (Sqr(2) + 0.1)) * 2.2
      End If
    Next
  End If
  
  For i = 1 To lngGla
    objProgCtrl.AddBar lngDivCount:=lngGla + 1
    rpdDPoint1.x = .dblCenterX - .dblMaxLength / 2 + _
                   .varElement(strGName(i), "POS") + _
                   .varElement(strGName(i), "PPX")
    rpdDPoint1.y = .dblCenterY - _
                   .varElement(strGName(i), "PPY")
    rpdDPoint2.x = xx(i)
    rpdDPoint2.y = yy
    
    Set rpdBalloonData = New zwDrawCAD.BalloonData
    rpdBalloonData.Create doc:=rpdDoc, _
                          StartPoint:=rpdDPoint1, CenterPoint:=rpdDPoint2, _
                          dRadius:=hh * (Sqr(2) + 0.1), sText:=" "
    objRapidPrim.SetUp prmObject:=rpdBalloonData, lngLayerNo:=CLng(Alpha(13)), _
                       lngPenStyle:=zwStyleSolid, lngPenWidth:=zwWidthThin
    rpdDoc.DataBase.Add Prim:=rpdBalloonData
    Set rpdBalloonData = Nothing
    
    For n = 1 To 2
      Set rpdTextData = New zwDrawCAD.TextData
      rpdTextData.FaceName = strFont
      rpdTextData.Height = hh
      rpdTextData.Width = ww
      Select Case n
        Case 1
          rpdTextData.Text = Mid(Trim(Left(strGName(i), 3)) & Space(4), 1, 4)
          rpdTextData.BasePosition = 1
        Case 2
          If Not CBool(fCheck(5)) Then
            rpdTextData.Text = "- 00"
          Else
            rpdTextData.Text = "-   "
          End If
          rpdTextData.BasePosition = 7
      End Select
      rpdTextData.BasePoint.x = rpdDPoint2.x
      rpdTextData.BasePoint.y = rpdDPoint2.y
      objRapidPrim.SetUp prmObject:=rpdTextData, lngLayerNo:=CLng(Alpha(13)), _
                         lngPenStyle:=zwStyleSolid, lngPenWidth:=zwWidthThin
      rpdDoc.DataBase.Add Prim:=rpdTextData
      Set rpdTextData = Nothing
    Next
    
    If Left(.varElement(strGName(i), "BAL"), 1) = "B" Then
      For n = 1 To 4
        Select Case n
          Case 1: tmpx = xx(i) - (hh * (Sqr(2) + 0.1)) * 1.1
          Case 2: tmpx = xx(i + 1) + (hh * (Sqr(2) + 0.1)) * 1.1
          Case Else: tmpx = (xx(i) + xx(i + 1)) / 2
        End Select
        tmpy = rpdDPoint2.y - hh * (Sqr(2) + 0.8)
        
        Set rpdArcData = New zwDrawCAD.ArcData
        rpdArcData.Radius = hh / 2
        rpdArcData.Center.x = tmpx + (hh / 2) * ff(n, 1)
        rpdArcData.Center.y = tmpy + (hh / 2) * ff(n, 2)
        rpdArcData.Start.x = tmpx + (hh / 2) * ff(n, 3)
        rpdArcData.Start.y = tmpy + (hh / 2) * ff(n, 4)
        rpdArcData.End.x = tmpx + (hh / 2) * ff(n, 5)
        rpdArcData.End.y = tmpy + (hh / 2) * ff(n, 6)
        objRapidPrim.SetUp prmObject:=rpdArcData, lngLayerNo:=CLng(Alpha(13)), _
                           lngPenStyle:=zwStyleSolid, lngPenWidth:=zwWidthThin
        rpdDoc.DataBase.Add Prim:=rpdArcData
        Set rpdArcData = Nothing
      Next
      
      For n = 1 To 2
        Select Case n
          Case 1
            rpdLineData.Start.x = xx(i) - (hh * (Sqr(2) + 0.1)) * 1.1 + hh / 2
            rpdLineData.End.x = (xx(i) + xx(i + 1)) / 2 - hh / 2
          Case 2
            rpdLineData.Start.x = xx(i + 1) + (hh * (Sqr(2) + 0.1)) * 1.1 - hh / 2
            rpdLineData.End.x = (xx(i) + xx(i + 1)) / 2 + hh / 2
        End Select
        rpdLineData.Start.y = rpdDPoint2.y - hh * (Sqr(2) + 0.8)
        rpdLineData.End.y = rpdDPoint2.y - hh * (Sqr(2) + 0.8)
        objRapidPrim.SetUp prmObject:=rpdLineData, lngLayerNo:=CLng(Alpha(13)), _
                           lngPenStyle:=zwStyleSolid, lngPenWidth:=zwWidthThin
        rpdDoc.DataBase.Add Prim:=rpdLineData
      Next
      
      Set rpdCircleData = New zwDrawCAD.CircleData
      rpdCircleData.Radius = hh * (Sqr(2) + 0.1)
      rpdCircleData.Center.x = (xx(i) + xx(i + 1)) / 2
      rpdCircleData.Center.y = rpdDPoint2.y - hh * (Sqr(2) * 2 + 1.7)
      objRapidPrim.SetUp prmObject:=rpdCircleData, lngLayerNo:=CLng(Alpha(13)), _
                         lngPenStyle:=zwStyleSolid, lngPenWidth:=zwWidthThin
      rpdDoc.DataBase.Add Prim:=rpdCircleData
      Set rpdCircleData = Nothing
      
      For n = 1 To 2
        Set rpdTextData = New zwDrawCAD.TextData
        rpdTextData.FaceName = strFont
        rpdTextData.Height = hh
        rpdTextData.Width = ww
        Select Case n
          Case 1
            rpdTextData.Text = _
              Mid(Trim(Left(.varElement(strGName(i), "BAL"), 3)) & Space(4), 1, 4)
            rpdTextData.BasePosition = 1
          Case 2
            If Not CBool(fCheck(5)) Then
              rpdTextData.Text = "- 00"
            Else
              rpdTextData.Text = "-   "
            End If
            rpdTextData.BasePosition = 7
        End Select
        rpdTextData.BasePoint.x = (xx(i) + xx(i + 1)) / 2
        rpdTextData.BasePoint.y = rpdDPoint2.y - hh * (Sqr(2) * 2 + 1.7)
        objRapidPrim.SetUp prmObject:=rpdTextData, lngLayerNo:=CLng(Alpha(13)), _
                           lngPenStyle:=zwStyleSolid, lngPenWidth:=zwWidthThin
        rpdDoc.DataBase.Add Prim:=rpdTextData
        Set rpdTextData = Nothing
      Next
    End If
  Next
  
  objProgCtrl.AddBar lngDivCount:=lngGla + 1
  If InStr(Cells(3, 3).Value, ",") > 0 Then
    Set rpdCircleData = New zwDrawCAD.CircleData
    rpdCircleData.Radius = hh * (Sqr(2) + 0.1)
    rpdCircleData.Center.x = xx(lngGla) + (hh * (Sqr(2) + 0.1)) * 4.4
    rpdCircleData.Center.y = rpdDPoint2.y - hh * (Sqr(2) * 2 + 1.7)
    objRapidPrim.SetUp prmObject:=rpdCircleData, lngLayerNo:=CLng(Alpha(13)), _
                       lngPenStyle:=zwStyleSolid, lngPenWidth:=zwWidthThin
    rpdDoc.DataBase.Add Prim:=rpdCircleData
    Set rpdCircleData = Nothing
    
    For n = 1 To 3
      Set rpdTextData = New zwDrawCAD.TextData
      rpdTextData.FaceName = strFont
      rpdTextData.Height = hh
      rpdTextData.Width = ww
      Select Case n
        Case 1
          rpdTextData.Text = "(光学仕様書)"
          If InStr(Cells(4, 6).Value, ",") > 0 Then
            rpdTextData.Text = Trim(Mid(Cells(4, 6).Value, _
                                  InStr(Cells(4, 6).Value, ",") + 1))
          End If
          rpdTextData.BasePosition = 1
          rpdTextData.BasePoint.y = rpdDPoint2.y - hh * (Sqr(2) + 1.3)
        Case 2
          rpdTextData.Text = Trim(Mid(Cells(3, 3).Value, _
                                InStr(Cells(3, 3).Value, ",") + 1))
          rpdTextData.BasePosition = 1
          rpdTextData.BasePoint.y = rpdDPoint2.y - hh * (Sqr(2) * 2 + 1.7)
        Case 3
          If Not CBool(fCheck(5)) Then
            rpdTextData.Text = "- 00"
          Else
            rpdTextData.Text = "-   "
          End If
          rpdTextData.BasePosition = 7
          rpdTextData.BasePoint.y = rpdDPoint2.y - hh * (Sqr(2) * 2 + 1.7)
      End Select
      rpdTextData.BasePoint.x = xx(lngGla) + (hh * (Sqr(2) + 0.1)) * 4.4
      objRapidPrim.SetUp prmObject:=rpdTextData, lngLayerNo:=CLng(Alpha(13)), _
                         lngPenStyle:=zwStyleSolid, lngPenWidth:=zwWidthThin
      rpdDoc.DataBase.Add Prim:=rpdTextData
      Set rpdTextData = Nothing
    Next
  End If
End With
Set rpdLineData = Nothing
Set rpdDPoint1 = Nothing
Set rpdDPoint2 = Nothing
Set objRapidPrim = Nothing

End Sub


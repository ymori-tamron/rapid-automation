VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsRapidGeo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False


Option Explicit
Option Base 1
Option Compare Binary


Public Sub AddG(ByRef gpdObject As clsGlassPlotData)

Dim rpdLineData As zwDrawCAD.LineData
Dim rpdPolyData As zwDrawCAD.PolyLineData
Dim rpdTextData As zwDrawCAD.TextData
Dim rpdDPoint1 As zwDrawCAD.DPoint
Dim rpdDPoint2 As zwDrawCAD.DPoint
Dim objRapidPrim As clsRapidPrim
Dim objRapidParam As clsRapidParam

Dim i As Long, j As Long, n As Long
Dim rc As Long, cc As Long, maxCol As Long
Dim dr As Long, dtm As Long, tr(2) As Long
Dim aa As Double, yy As Double, zz As Double
Dim hh As Double, ww As Double
Dim strFont As String, txt(2) As String

Set objRapidParam = New clsRapidParam
objRapidParam.Import strConfig:="Param"
With objRapidParam
  strFont = .strItem(1)
  aa = .dblXY(1, 1) / rpdDoc.DrawInfo.Numerator
  
  rc = 0
  For i = 1 To .lngMaxCount
    If .strItem(i) = "äÙâΩåˆç∑" Then
      yy = .dblXY(i, 2) / rpdDoc.DrawInfo.Numerator
      hh = .dblHW(i, 1) / rpdDoc.DrawInfo.Numerator
      ww = .dblHW(i, 2) / rpdDoc.DrawInfo.Numerator
      dr = .lngVal(i, 1)
      tr(1) = .lngVal(i + 1, 1)
      tr(2) = .lngVal(i + 2, 1)
      rc = i: Exit For
    End If
  Next
End With
Set objRapidParam = Nothing

If rc = 0 Then
  objProgCtrl.AddBar lngDivCount:=1
  Exit Sub
End If

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
    If InStr(ActiveSheet.Name, "ASP") = 0 Then
      If Not IsEmpty(.Cells(dr, cc).Value) Then
        dtm = .Cells(dr, cc).Value
        txt(1) = .Cells(tr(1), cc + 1).Text
        txt(2) = .Cells(tr(2), cc + 1).Text
      Else
        objProgCtrl.AddBar lngDivCount:=1
        Exit Sub
      End If
    Else
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
With gpdObject
  objProgCtrl.AddBar lngDivCount:=2
  If txt(1) <> "" Then
    rpdDPoint1.x = .dblCenterX(dtm) + .dblCornerZ(1, dtm)
    rpdDPoint1.y = .dblCenterY(dtm) + .dblCornerH(2, dtm) + yy
    rpdDPoint2.x = .dblCenterX(dtm) + .dblCornerZ(1, dtm) + (aa * 2) * (dtm * 2 - 3)
    rpdDPoint2.y = .dblCenterY(dtm) + .dblCornerH(2, dtm) + yy
    
    Set rpdPolyData = New zwDrawCAD.PolyLineData
    rpdPolyData.CreateArrow Type:=zwMarkArrow, _
                            StartPoint:=rpdDPoint1, EndPoint:=rpdDPoint2, _
                            dSize:=aa
    objRapidPrim.SetUp prmObject:=rpdPolyData, lngLayerNo:=CLng(Alpha(13)), _
                       lngPenStyle:=zwStyleSolid, lngPenWidth:=zwWidthThin
    rpdDoc.DataBase.Add Prim:=rpdPolyData
    Set rpdPolyData = Nothing
    
    If dtm = 1 Then
      rpdDPoint1.x = .dblCenterX(1) + .dblCornerZ(1, 1) - aa * 2
      rpdDPoint2.x = .dblCenterX(1) + .dblCornerZ(1, 1) - aa * 2
    Else
      rpdDPoint1.x = .dblCenterX(2) + .dblCornerZ(1, 2) + aa * 2 + ww * 6.6
      rpdDPoint2.x = .dblCenterX(2) + .dblCornerZ(1, 2) + aa * 2 + ww * 6.6
    End If
    rpdDPoint1.y = .dblCenterY(dtm) + .dblCornerH(2, dtm) + yy - hh * 0.6
    rpdDPoint2.y = .dblCenterY(dtm) + .dblCornerH(2, dtm) + yy + hh * 0.6
    
    For n = 1 To 7
      Select Case n
        Case 1
          rpdLineData.Start.x = .dblCenterX(dtm) + .dblCornerZ(1, dtm)
          rpdLineData.Start.y = .dblCenterY(dtm) + .dblCornerH(1, dtm) + aa / 2
          rpdLineData.End.x = .dblCenterX(dtm) + .dblCornerZ(1, dtm)
          rpdLineData.End.y = .dblCenterY(dtm) + .dblCornerH(2, dtm) + yy + aa / 2
        Case 2
          rpdLineData.Start.x = rpdDPoint1.x
          rpdLineData.Start.y = rpdDPoint1.y
          rpdLineData.End.x = rpdDPoint2.x
          rpdLineData.End.y = rpdDPoint2.y
        Case 3
          rpdLineData.Start.x = rpdLineData.Start.x - ww * 2.4
          rpdLineData.End.x = rpdLineData.End.x - ww * 2.4
        Case 4
          rpdLineData.Start.x = rpdLineData.Start.x - ww * 3
          rpdLineData.End.x = rpdLineData.End.x - ww * 3
        Case 5
          rpdLineData.Start.x = rpdLineData.Start.x - ww * 1.2
          rpdLineData.End.x = rpdLineData.End.x - ww * 1.2
        Case 6
          rpdLineData.End.x = rpdDPoint1.x
          rpdLineData.End.y = rpdDPoint1.y
        Case 7
          rpdLineData.Start.y = rpdDPoint2.y
          rpdLineData.End.y = rpdDPoint2.y
      End Select
      objRapidPrim.SetUp prmObject:=rpdLineData, lngLayerNo:=CLng(Alpha(13)), _
                         lngPenStyle:=zwStyleSolid, lngPenWidth:=zwWidthThin
      rpdDoc.DataBase.Add Prim:=rpdLineData
    Next
    
    For n = 1 To 3
      Set rpdTextData = New zwDrawCAD.TextData
      rpdTextData.FaceName = strFont
      rpdTextData.Height = hh
      rpdTextData.Width = ww
      rpdTextData.BasePosition = 4
      Select Case n
        Case 1
          rpdTextData.Text = "åıé≤"
          rpdTextData.BasePoint.x = rpdDPoint2.x - ww * 1.2
        Case 2
          rpdTextData.Text = txt(1)
          rpdTextData.BasePoint.x = rpdDPoint2.x - ww * 3.9
        Case 3
          rpdTextData.Text = "Å€"
          rpdTextData.BasePoint.x = rpdDPoint2.x - ww * 6
      End Select
      rpdTextData.BasePoint.y = .dblCenterY(dtm) + .dblCornerH(2, dtm) + yy
      objRapidPrim.SetUp prmObject:=rpdTextData, lngLayerNo:=CLng(Alpha(13)), _
                         lngPenStyle:=zwStyleSolid, lngPenWidth:=zwWidthThin
      rpdDoc.DataBase.Add Prim:=rpdTextData
      Set rpdTextData = Nothing
    Next
  End If
  
  
  objProgCtrl.AddBar lngDivCount:=2
  If txt(2) <> "" Then
    For j = 1 To 2
      rpdDPoint1.x = .dblCenterX(j) + .dblCornerZ(1, j)
      rpdDPoint1.y = .dblCenterY(j) + .dblCornerH(2, j) + yy + hh * 1.2
      rpdDPoint2.x = .dblCenterX(j) + .dblCornerZ(1, j) + (aa * 2) * (j * 2 - 3)
      rpdDPoint2.y = .dblCenterY(j) + .dblCornerH(2, j) + yy + hh * 1.2
      
      Set rpdPolyData = New zwDrawCAD.PolyLineData
      rpdPolyData.CreateArrow Type:=zwMarkArrow, _
                              StartPoint:=rpdDPoint1, EndPoint:=rpdDPoint2, _
                              dSize:=aa
      objRapidPrim.SetUp prmObject:=rpdPolyData, lngLayerNo:=CLng(Alpha(13)), _
                         lngPenStyle:=zwStyleSolid, lngPenWidth:=zwWidthThin
      rpdDoc.DataBase.Add Prim:=rpdPolyData
      Set rpdPolyData = Nothing
      
      zz = ww * 2.1 - (.dblCenterX(2) + .dblCornerZ(1, 2) - _
                       .dblCenterX(1) - .dblCornerZ(1, 1)) / 2
      If zz < 0 Then
        zz = 0
      Else
        rpdLineData.Start.x = .dblCenterX(j) + .dblCornerZ(1, j) + (aa * 2) * (j * 2 - 3)
        rpdLineData.Start.y = .dblCenterY(j) + .dblCornerH(2, j) + yy + hh * 1.2
        rpdLineData.End.x = .dblCenterX(j) + .dblCornerZ(1, j) + (aa * 2 + zz) * (j * 2 - 3)
        rpdLineData.End.y = .dblCenterY(j) + .dblCornerH(2, j) + yy + hh * 1.2
        objRapidPrim.SetUp prmObject:=rpdLineData, lngLayerNo:=CLng(Alpha(13)), _
                           lngPenStyle:=zwStyleSolid, lngPenWidth:=zwWidthThin
        rpdDoc.DataBase.Add Prim:=rpdLineData
      End If
      
      For n = 1 To 3
        Select Case n
          Case 1
            rpdLineData.Start.x = .dblCenterX(j) + .dblCornerZ(1, j)
            rpdLineData.Start.y = .dblCenterY(j) + .dblCornerH(1, j) + aa / 2
            rpdLineData.End.x = .dblCenterX(j) + .dblCornerZ(1, j)
            rpdLineData.End.y = .dblCenterY(j) + .dblCornerH(2, j) + (yy + hh * 1.2) + aa / 2
          Case 2
            rpdLineData.Start.x = .dblCenterX(j) + .dblCornerZ(1, j) + (aa * 2 + zz) * (j * 2 - 3)
            rpdLineData.Start.y = .dblCenterY(j) + .dblCornerH(2, j) + yy + hh * 1.2
            rpdLineData.End.x = .dblCenterX(j) + .dblCornerZ(1, j) + (aa * 2 + zz) * (j * 2 - 3)
            rpdLineData.End.y = .dblCenterY(j) + .dblCornerH(2, j) + yy + hh * 3
          Case 3
            rpdLineData.Start.x = .dblCenterX(j) + .dblCornerZ(1, j) + (aa * 2 + zz) * (j * 2 - 3)
            rpdLineData.Start.y = .dblCenterY(j) + .dblCornerH(2, j) + yy + hh * 3
            rpdLineData.End.x = (.dblCenterX(1) + .dblCornerZ(1, 1) + _
                                 .dblCenterX(2) + .dblCornerZ(1, 2)) / 2 + (ww * 2.1) * (j * 2 - 3)
            rpdLineData.End.y = .dblCenterY(j) + .dblCornerH(2, j) + yy + hh * 3
        End Select
        objRapidPrim.SetUp prmObject:=rpdLineData, lngLayerNo:=CLng(Alpha(13)), _
                           lngPenStyle:=zwStyleSolid, lngPenWidth:=zwWidthThin
        rpdDoc.DataBase.Add Prim:=rpdLineData
      Next
    Next
    
    rpdDPoint1.x = (.dblCenterX(1) + .dblCornerZ(1, 1) + _
                    .dblCenterX(2) + .dblCornerZ(1, 2)) / 2 + ww * 2.1
    rpdDPoint1.y = .dblCenterY(dtm) + .dblCornerH(2, dtm) + (yy + hh * 3) - hh * 0.6
    rpdDPoint2.x = (.dblCenterX(1) + .dblCornerZ(1, 1) + _
                    .dblCenterX(2) + .dblCornerZ(1, 2)) / 2 + ww * 2.1
    rpdDPoint2.y = .dblCenterY(dtm) + .dblCornerH(2, dtm) + (yy + hh * 3) + hh * 0.6
    
    For n = 1 To 5
      Select Case n
        Case 1
          rpdLineData.Start.x = rpdDPoint1.x
          rpdLineData.Start.y = rpdDPoint1.y
          rpdLineData.End.x = rpdDPoint2.x
          rpdLineData.End.y = rpdDPoint2.y
        Case 2
          rpdLineData.Start.x = rpdLineData.Start.x - ww * 3
          rpdLineData.End.x = rpdLineData.End.x - ww * 3
        Case 3
          rpdLineData.Start.x = rpdLineData.Start.x - ww * 1.2
          rpdLineData.End.x = rpdLineData.End.x - ww * 1.2
        Case 4
          rpdLineData.End.x = rpdDPoint1.x
          rpdLineData.End.y = rpdDPoint1.y
        Case 5
          rpdLineData.Start.y = rpdDPoint2.y
          rpdLineData.End.y = rpdDPoint2.y
      End Select
      objRapidPrim.SetUp prmObject:=rpdLineData, lngLayerNo:=CLng(Alpha(13)), _
                         lngPenStyle:=zwStyleSolid, lngPenWidth:=zwWidthThin
      rpdDoc.DataBase.Add Prim:=rpdLineData
    Next
    
    For n = 1 To 2
      Set rpdTextData = New zwDrawCAD.TextData
      rpdTextData.FaceName = strFont
      rpdTextData.Height = hh
      rpdTextData.Width = ww
      rpdTextData.BasePosition = 4
      Select Case n
        Case 1
          rpdTextData.Text = txt(2)
          rpdTextData.BasePoint.x = rpdDPoint2.x - ww * 1.5
        Case 2
          rpdTextData.Text = "Åa"
          rpdTextData.BasePoint.x = rpdDPoint2.x - ww * 3.6
      End Select
      rpdTextData.BasePoint.y = .dblCenterY(dtm) + .dblCornerH(2, dtm) + yy + hh * 3
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


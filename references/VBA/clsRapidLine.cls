VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsRapidLine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False


Option Explicit
Option Base 1
Option Compare Binary


Public Sub AddG(ByRef gpdObject As clsGlassPlotData)

Dim rpdLineData As zwDrawCAD.LineData
Dim rpdArcData As zwDrawCAD.ArcData
Dim objRapidPrim As clsRapidPrim

Dim i As Long, j As Long, m As Long

Set rpdLineData = New zwDrawCAD.LineData
Set objRapidPrim = New clsRapidPrim
With gpdObject
  For m = -1 To 1 Step 2
    For j = 1 To 2
      If .blnAsp(j) Then
        For i = 0 To .lngMaxCount - 1
          objProgCtrl.AddBar lngDivCount:=.lngMaxCount * 4
          rpdLineData.Start.x = .dblCenterX(j) + .dblSagZ(i, j)
          rpdLineData.Start.y = .dblCenterY(j) + .dblDeltaH(j) * i * m
          rpdLineData.End.x = .dblCenterX(j) + .dblSagZ(i + 1, j)
          rpdLineData.End.y = .dblCenterY(j) + .dblDeltaH(j) * (i + 1) * m
          objRapidPrim.SetUp prmObject:=rpdLineData, lngLayerNo:=CLng(Alpha(10)), _
                             lngPenStyle:=zwStyleSolid, lngPenWidth:=zwWidthMidThin
          rpdDoc.DataBase.Add Prim:=rpdLineData
        Next
      Else
        objProgCtrl.AddBar lngDivCount:=4
        Set rpdArcData = New zwDrawCAD.ArcData
        rpdArcData.Radius = Abs(.dblRadius(j))
        rpdArcData.Center.x = .dblCenterX(j) + .dblRadius(j)
        rpdArcData.Center.y = .dblCenterY(j)
        If .dblRadius(j) * m > 0 Then
          rpdArcData.Start.x = .dblCenterX(j) + .dblCornerZ(1, j)
          rpdArcData.Start.y = .dblCenterY(j) + .dblDeltaH(j) * .lngMaxCount * m
          rpdArcData.End.x = .dblCenterX(j)
          rpdArcData.End.y = .dblCenterY(j)
        Else
          rpdArcData.Start.x = .dblCenterX(j)
          rpdArcData.Start.y = .dblCenterY(j)
          rpdArcData.End.x = .dblCenterX(j) + .dblCornerZ(1, j)
          rpdArcData.End.y = .dblCenterY(j) + .dblDeltaH(j) * .lngMaxCount * m
        End If
        objRapidPrim.SetUp prmObject:=rpdArcData, lngLayerNo:=CLng(Alpha(10)), _
                           lngPenStyle:=zwStyleSolid, lngPenWidth:=zwWidthMidThin
        rpdDoc.DataBase.Add Prim:=rpdArcData
        Set rpdArcData = Nothing
      End If
      
      If .dblSurfaceH(j) <> 0 Then
        If Left(ActiveSheet.Name, 1) = "G" Or _
           Left(ActiveSheet.Name, 1) = "H" And _
           Not .blnAsp(j) Then
          rpdLineData.Start.x = .dblCenterX(j) + .dblCornerZ(1, j)
          rpdLineData.Start.y = .dblCenterY(j) + .dblSurfaceH(j) * m
          rpdLineData.End.x = .dblCenterX(j) + .dblCornerZ(1, j)
          rpdLineData.End.y = .dblCenterY(j) + .dblCornerH(1, j) * m
          objRapidPrim.SetUp prmObject:=rpdLineData, lngLayerNo:=CLng(Alpha(10)), _
                             lngPenStyle:=zwStyleSolid, lngPenWidth:=zwWidthMidThin
          rpdDoc.DataBase.Add Prim:=rpdLineData
        End If
      End If
      
      If .dblCornerH(2, j) - .dblCornerH(1, j) > 0 And _
         Left(ActiveSheet.Name, 1) = "G" Then
        rpdLineData.Start.x = .dblCenterX(j) + .dblCornerZ(1, j)
        rpdLineData.Start.y = .dblCenterY(j) + .dblCornerH(1, j) * m
        rpdLineData.End.x = .dblCenterX(j) + .dblCornerZ(2, j)
        rpdLineData.End.y = .dblCenterY(j) + .dblCornerH(2, j) * m
        objRapidPrim.SetUp prmObject:=rpdLineData, lngLayerNo:=CLng(Alpha(10)), _
                           lngPenStyle:=zwStyleSolid, lngPenWidth:=zwWidthMidThin
        rpdDoc.DataBase.Add Prim:=rpdLineData
      End If
    Next
    
    If Left(ActiveSheet.Name, 1) = "H" Then
      rpdLineData.Start.x = .dblCenterX(1) + .dblCornerZ(1, 1)
      rpdLineData.End.x = .dblCenterX(2) + .dblCornerZ(1, 2)
      If .blnAsp(1) Then
        rpdLineData.Start.y = .dblCenterY(1) + .dblSurfaceH(1) * m
        rpdLineData.End.y = .dblCenterY(2) + .dblCornerH(2, 2) * m
      Else
        rpdLineData.Start.y = .dblCenterY(1) + .dblCornerH(2, 1) * m
        rpdLineData.End.y = .dblCenterY(2) + .dblSurfaceH(2) * m
      End If
    Else
      rpdLineData.Start.x = .dblCenterX(1) + .dblCornerZ(2, 1)
      rpdLineData.Start.y = .dblCenterY(1) + .dblCornerH(2, 1) * m
      rpdLineData.End.x = .dblCenterX(2) + .dblCornerZ(2, 2)
      rpdLineData.End.y = .dblCenterY(2) + .dblCornerH(2, 2) * m
    End If
    objRapidPrim.SetUp prmObject:=rpdLineData, lngLayerNo:=CLng(Alpha(10)), _
                       lngPenStyle:=zwStyleSolid, lngPenWidth:=zwWidthMidThin
    rpdDoc.DataBase.Add Prim:=rpdLineData
  Next
  
  If Left(ActiveSheet.Name, 1) <> "H" Then
    rpdLineData.Start.x = .dblCenterX(1) - 5 / rpdDoc.DrawInfo.Numerator
    rpdLineData.Start.y = .dblCenterY(1)
    rpdLineData.End.x = .dblCenterX(2) + 5 / rpdDoc.DrawInfo.Numerator
    rpdLineData.End.y = .dblCenterY(2)
    objRapidPrim.SetUp prmObject:=rpdLineData, lngLayerNo:=CLng(Alpha(11)), _
                       lngPenStyle:=zwStyleS1DashDot, lngPenWidth:=zwWidthThin
    rpdDoc.DataBase.Add Prim:=rpdLineData
  End If
End With
Set rpdLineData = Nothing
Set objRapidPrim = Nothing

End Sub


Public Sub AddP(ByRef ppdObject As clsPressPlotData)

Dim rpdLineData As zwDrawCAD.LineData
Dim rpdArcData As zwDrawCAD.ArcData
Dim objRapidPrim As clsRapidPrim

Dim j As Long, m As Long

Set rpdLineData = New zwDrawCAD.LineData
Set objRapidPrim = New clsRapidPrim
With ppdObject
  For m = -1 To 1 Step 2
    For j = 1 To 2
      objProgCtrl.AddBar lngDivCount:=4
      Set rpdArcData = New zwDrawCAD.ArcData
      rpdArcData.Radius = Abs(.dblRadius(j))
      rpdArcData.Center.x = .dblCenterX(j) + .dblRadius(j)
      rpdArcData.Center.y = .dblCenterY(j)
      If .dblRadius(j) * m > 0 Then
        rpdArcData.Start.x = .dblCenterX(j) + .dblCornerZ(1, j)
        rpdArcData.Start.y = .dblCenterY(j) + .dblSurfaceH(j) * m
        rpdArcData.End.x = .dblCenterX(j)
        rpdArcData.End.y = .dblCenterY(j)
      Else
        rpdArcData.Start.x = .dblCenterX(j)
        rpdArcData.Start.y = .dblCenterY(j)
        rpdArcData.End.x = .dblCenterX(j) + .dblCornerZ(1, j)
        rpdArcData.End.y = .dblCenterY(j) + .dblSurfaceH(j) * m
      End If
      objRapidPrim.SetUp prmObject:=rpdArcData, lngLayerNo:=lngPre(1), _
                         lngPenStyle:=zwStyleSolid, lngPenWidth:=zwWidthMidThin
      rpdDoc.DataBase.Add Prim:=rpdArcData
      Set rpdArcData = Nothing
      
      If .dblCornerH(1, j) - .dblSurfaceH(j) > 0 Then
        rpdLineData.Start.x = .dblCenterX(j) + .dblCornerZ(1, j)
        rpdLineData.Start.y = .dblCenterY(j) + .dblSurfaceH(j) * m
        rpdLineData.End.x = .dblCenterX(j) + .dblCornerZ(1, j)
        rpdLineData.End.y = .dblCenterY(j) + .dblCornerH(1, j) * m
        objRapidPrim.SetUp prmObject:=rpdLineData, lngLayerNo:=lngPre(1), _
                           lngPenStyle:=zwStyleSolid, lngPenWidth:=zwWidthMidThin
        rpdDoc.DataBase.Add Prim:=rpdLineData
      End If
      
      If .dblCornerH(2, j) - .dblCornerH(1, j) > 0 Then
        rpdLineData.Start.x = .dblCenterX(j) + .dblCornerZ(1, j)
        rpdLineData.Start.y = .dblCenterY(j) + .dblCornerH(1, j) * m
        rpdLineData.End.x = .dblCenterX(j) + .dblCornerZ(2, j)
        rpdLineData.End.y = .dblCenterY(j) + .dblCornerH(2, j) * m
        objRapidPrim.SetUp prmObject:=rpdLineData, lngLayerNo:=lngPre(1), _
                           lngPenStyle:=zwStyleSolid, lngPenWidth:=zwWidthMidThin
        rpdDoc.DataBase.Add Prim:=rpdLineData
      End If
    Next
    
    rpdLineData.Start.x = .dblCenterX(1) + .dblCornerZ(2, 1)
    rpdLineData.Start.y = .dblCenterY(1) + .dblCornerH(2, 1) * m
    rpdLineData.End.x = .dblCenterX(2) + .dblCornerZ(2, 2)
    rpdLineData.End.y = .dblCenterY(2) + .dblCornerH(2, 2) * m
    objRapidPrim.SetUp prmObject:=rpdLineData, lngLayerNo:=lngPre(1), _
                       lngPenStyle:=zwStyleSolid, lngPenWidth:=zwWidthMidThin
    rpdDoc.DataBase.Add Prim:=rpdLineData
  Next
  
  rpdLineData.Start.x = .dblCenterX(1) - 5 / rpdDoc.DrawInfo.Numerator
  rpdLineData.Start.y = .dblCenterY(1)
  rpdLineData.End.x = .dblCenterX(2) + 5 / rpdDoc.DrawInfo.Numerator
  rpdLineData.End.y = .dblCenterY(2)
  objRapidPrim.SetUp prmObject:=rpdLineData, lngLayerNo:=lngPre(2), _
                     lngPenStyle:=zwStyleS1DashDot, lngPenWidth:=zwWidthThin
  rpdDoc.DataBase.Add Prim:=rpdLineData
End With
Set rpdLineData = Nothing
Set objRapidPrim = Nothing

End Sub


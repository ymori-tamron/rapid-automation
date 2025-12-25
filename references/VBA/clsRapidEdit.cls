VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsRapidEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False


Option Explicit
Option Base 1
Option Compare Binary


Public Sub EditAspFrame(ByRef gpdObject As clsGlassPlotData)

Dim rpdDPoint1 As zwDrawCAD.DPoint
Dim rpdDPoint2 As zwDrawCAD.DPoint

Dim xx(2) As Double

If InStr(ActiveSheet.Name, "ASP") > 0 Then
  If CBool(fCheck(4)) Then
    xx(1) = 201: xx(2) = 367
  Else
    xx(1) = 196: xx(2) = 404
  End If
  
  If Not gpdObject.blnAsp(1) Then
    Set rpdDPoint1 = New zwDrawCAD.DPoint
    Set rpdDPoint2 = New zwDrawCAD.DPoint
    With rpdDoc
      rpdDPoint1.x = xx(1) / .DrawInfo.Numerator
      rpdDPoint2.x = xx(2) / .DrawInfo.Numerator
      
      rpdDPoint1.y = 220 / .DrawInfo.Numerator
      rpdDPoint2.y = 166 / .DrawInfo.Numerator
      .DataBase.SelectArea _
        FirstPoint:=rpdDPoint1, LastPoint:=rpdDPoint2, bSelect:=True
      .DataBase.RemoveSelect
      
      rpdDPoint1.y = 156 / .DrawInfo.Numerator
      rpdDPoint2.y = 102 / .DrawInfo.Numerator
      .DataBase.SelectArea _
        FirstPoint:=rpdDPoint1, LastPoint:=rpdDPoint2, bSelect:=True
      .DataBase.MoveS _
        dX:=0, dY:=48.969 / .DrawInfo.Numerator, KeepSelect:=zwKeepSelectNon
      .Views(0).Redraw
    End With
    Set rpdDPoint1 = Nothing
    Set rpdDPoint2 = Nothing
  End If
  
  If Not gpdObject.blnAsp(2) Then
    Set rpdDPoint1 = New zwDrawCAD.DPoint
    Set rpdDPoint2 = New zwDrawCAD.DPoint
    With rpdDoc
      rpdDPoint1.x = xx(1) / .DrawInfo.Numerator
      rpdDPoint2.x = xx(2) / .DrawInfo.Numerator
      
      rpdDPoint1.y = 156 / .DrawInfo.Numerator
      rpdDPoint2.y = 102 / .DrawInfo.Numerator
      .DataBase.SelectArea _
        FirstPoint:=rpdDPoint1, LastPoint:=rpdDPoint2, bSelect:=True
      .DataBase.RemoveSelect
      
      rpdDPoint1.y = 220 / .DrawInfo.Numerator
      rpdDPoint2.y = 166 / .DrawInfo.Numerator
      .DataBase.SelectArea _
        FirstPoint:=rpdDPoint1, LastPoint:=rpdDPoint2, bSelect:=True
      .DataBase.MoveS _
        dX:=0, dY:=-15 / .DrawInfo.Numerator, KeepSelect:=zwKeepSelectNon
      .Views(0).Redraw
    End With
    Set rpdDPoint1 = Nothing
    Set rpdDPoint2 = Nothing
  End If
End If

End Sub


Public Sub EditGlaDBBlank()

Dim rpdLineData As zwDrawCAD.LineData
Dim objRapidPrim As clsRapidPrim

Dim i As Long, j As Long

If InStr(ActiveSheet.Name, "ASP") = 0 Then
  Set rpdLineData = New zwDrawCAD.LineData
  Set objRapidPrim = New clsRapidPrim
  With rpdLineData
    .Start.x = 152.54 / rpdDoc.DrawInfo.Numerator
    .End.x = 190.54 / rpdDoc.DrawInfo.Numerator
    For i = 1 To 6
      If IsEmpty(Cells(i + 30, 6).Value) Then
        For j = 1 To 2
          Select Case j
            Case 1
              .Start.y = (267.88 - 5 * (i - 1)) / rpdDoc.DrawInfo.Numerator
              .End.y = (262.88 - 5 * (i - 1)) / rpdDoc.DrawInfo.Numerator
            Case 2
              .Start.y = (262.88 - 5 * (i - 1)) / rpdDoc.DrawInfo.Numerator
              .End.y = (267.88 - 5 * (i - 1)) / rpdDoc.DrawInfo.Numerator
          End Select
          objRapidPrim.SetUp prmObject:=rpdLineData, lngLayerNo:=CLng(Alpha(12)), _
                             lngPenStyle:=zwStyleSolid, lngPenWidth:=zwWidthThin
          rpdDoc.DataBase.Add Prim:=rpdLineData
        Next
      End If
    Next
  End With
  Set rpdLineData = Nothing
  Set objRapidPrim = Nothing
End If

End Sub


Public Sub RemoveGmNote()

Dim rpdDPoint1 As zwDrawCAD.DPoint
Dim rpdDPoint2 As zwDrawCAD.DPoint

Set rpdDPoint1 = New zwDrawCAD.DPoint
Set rpdDPoint2 = New zwDrawCAD.DPoint
With rpdDoc
  rpdDPoint1.x = 35 / .DrawInfo.Numerator
  rpdDPoint2.x = 210 / .DrawInfo.Numerator
  
  rpdDPoint1.y = 53.5 / .DrawInfo.Numerator
  rpdDPoint2.y = 50.5 / .DrawInfo.Numerator
  .DataBase.SelectArea _
    FirstPoint:=rpdDPoint1, LastPoint:=rpdDPoint2, bSelect:=True
  
  rpdDPoint1.y = 48.5 / .DrawInfo.Numerator
  rpdDPoint2.y = 45.5 / .DrawInfo.Numerator
  .DataBase.SelectArea _
    FirstPoint:=rpdDPoint1, LastPoint:=rpdDPoint2, bSelect:=True
  
  rpdDPoint1.y = 43.5 / .DrawInfo.Numerator
  rpdDPoint2.y = 40.5 / .DrawInfo.Numerator
  .DataBase.SelectArea _
    FirstPoint:=rpdDPoint1, LastPoint:=rpdDPoint2, bSelect:=True
  
  rpdDPoint1.x = 318.1 / .DrawInfo.Numerator
  rpdDPoint2.x = 342.4 / .DrawInfo.Numerator
  
  rpdDPoint1.y = 32.5 / .DrawInfo.Numerator
  rpdDPoint2.y = 26.5 / .DrawInfo.Numerator
  .DataBase.SelectArea _
    FirstPoint:=rpdDPoint1, LastPoint:=rpdDPoint2, bSelect:=True
  
  .DataBase.RemoveSelect
  .Views(0).Redraw
End With
Set rpdDPoint1 = Nothing
Set rpdDPoint2 = Nothing

End Sub


Public Sub RemoveExStd()

Dim rpdDPoint1 As zwDrawCAD.DPoint
Dim rpdDPoint2 As zwDrawCAD.DPoint

If lngPre(3) <> CLng(Alpha(12)) Then
  With rpdDoc.DrawLayers
    .Item(.GetLayer(lNumber:=CLng(Alpha(12)))).Visible = False
    .Item(.GetLayer(lNumber:=lngPre(3))).Visible = True
  End With
End If

Set rpdDPoint1 = New zwDrawCAD.DPoint
Set rpdDPoint2 = New zwDrawCAD.DPoint
With rpdDoc
  rpdDPoint1.x = 104.7 / .DrawInfo.Numerator
  rpdDPoint2.x = 129.1 / .DrawInfo.Numerator
  
  rpdDPoint1.y = 26.9 / .DrawInfo.Numerator
  rpdDPoint2.y = 20.9 / .DrawInfo.Numerator
  .DataBase.SelectArea _
    FirstPoint:=rpdDPoint1, LastPoint:=rpdDPoint2, bSelect:=True
  
  .DataBase.RemoveSelect
  .Views(0).Redraw
End With
Set rpdDPoint1 = Nothing
Set rpdDPoint2 = Nothing

If lngPre(3) <> CLng(Alpha(12)) Then
  With rpdDoc.DrawLayers
    .Item(.GetLayer(lNumber:=CLng(Alpha(12)))).Visible = True
    .Item(.GetLayer(lNumber:=lngPre(3))).Visible = False
  End With
End If

End Sub


Public Sub RemoveNoteFrame()

Dim rpdDPoint1 As zwDrawCAD.DPoint
Dim rpdDPoint2 As zwDrawCAD.DPoint

Set rpdDPoint1 = New zwDrawCAD.DPoint
Set rpdDPoint2 = New zwDrawCAD.DPoint
With rpdDoc
  rpdDPoint1.x = 20 / .DrawInfo.Numerator
  rpdDPoint2.x = 146.7 / .DrawInfo.Numerator
  
  rpdDPoint1.y = 276.9 / .DrawInfo.Numerator
  rpdDPoint2.y = 222.4 / .DrawInfo.Numerator
  .DataBase.SelectArea _
    FirstPoint:=rpdDPoint1, LastPoint:=rpdDPoint2, bSelect:=True
  
  .DataBase.RemoveSelect
  .Views(0).Redraw
End With
Set rpdDPoint1 = Nothing
Set rpdDPoint2 = Nothing

End Sub


Public Sub RemoveGlaDBFrame()

Dim rpdDPoint1 As zwDrawCAD.DPoint
Dim rpdDPoint2 As zwDrawCAD.DPoint

Set rpdDPoint1 = New zwDrawCAD.DPoint
Set rpdDPoint2 = New zwDrawCAD.DPoint
With rpdDoc
  rpdDPoint1.x = 152 / .DrawInfo.Numerator
  rpdDPoint2.x = 191 / .DrawInfo.Numerator
  
  rpdDPoint1.y = 273.4 / .DrawInfo.Numerator
  rpdDPoint2.y = 227.9 / .DrawInfo.Numerator
  .DataBase.SelectArea _
    FirstPoint:=rpdDPoint1, LastPoint:=rpdDPoint2, bSelect:=True
  
  .DataBase.RemoveSelect
  .Views(0).Redraw
End With
Set rpdDPoint1 = Nothing
Set rpdDPoint2 = Nothing

End Sub


Public Sub CreateHybBase(ByRef gpdObject As clsGlassPlotData)

Dim rpdLineData As zwDrawCAD.LineData
Dim objRapidPrim As clsRapidPrim
Dim objRapidLine As clsRapidLine
Dim objGPData As clsGlassPlotData

Dim i As Long, bc As Long, ac As Long
Dim asName As String

asName = ActiveSheet.Name

Set rpdLineData = New zwDrawCAD.LineData
Set objRapidPrim = New clsRapidPrim
Set objRapidLine = New clsRapidLine
Set objGPData = New clsGlassPlotData
With rpdDoc.DataBase
  bc = .Count
  Worksheets("G" & Mid(asName, 2, 2)).Activate
  objGPData.Calc
  objRapidLine.AddG gpdObject:=objGPData
  Worksheets(asName).Activate
  .Item(.Count - 1).Selected = True
  .RemoveSelect
  ac = .Count - 1
  For i = bc To ac
    With .Item(i)
      .Edit = True
      .LayerNo = CLng(Alpha(13))
      .PenStyle = zwStyleSDash
      .PenWidth = zwWidthThin
      .Edit = False
      .Selected = True
    End With
  Next
End With

With rpdDoc
  If gpdObject.blnAsp(1) Then
    .DataBase.MoveS _
      dX:=(objGPData.dblCenterX(2) - objGPData.dblCenterX(1) + _
           gpdObject.dblCenterX(2) - gpdObject.dblCenterX(1)) / 2, _
      dY:=0, _
      KeepSelect:=zwKeepSelectNon
    rpdLineData.Start.x = gpdObject.dblCenterX(1) - 5 / .DrawInfo.Numerator
    rpdLineData.Start.y = gpdObject.dblCenterY(1)
    rpdLineData.End.x = gpdObject.dblCenterX(2) + 5 / .DrawInfo.Numerator + _
                        objGPData.dblCenterX(2) - objGPData.dblCenterX(1)
    rpdLineData.End.y = gpdObject.dblCenterY(2)
  Else
    .DataBase.MoveS _
      dX:=(objGPData.dblCenterX(1) - objGPData.dblCenterX(2) + _
           gpdObject.dblCenterX(1) - gpdObject.dblCenterX(2)) / 2, _
      dY:=0, _
      KeepSelect:=zwKeepSelectNon
    rpdLineData.Start.x = gpdObject.dblCenterX(1) - 5 / .DrawInfo.Numerator + _
                          objGPData.dblCenterX(1) - objGPData.dblCenterX(2)
    rpdLineData.Start.y = gpdObject.dblCenterY(1)
    rpdLineData.End.x = gpdObject.dblCenterX(2) + 5 / .DrawInfo.Numerator
    rpdLineData.End.y = gpdObject.dblCenterY(2)
  End If
  .Views(0).Redraw
  objRapidPrim.SetUp prmObject:=rpdLineData, lngLayerNo:=CLng(Alpha(11)), _
                     lngPenStyle:=zwStyleS1DashDot, lngPenWidth:=zwWidthThin
  .DataBase.Add Prim:=rpdLineData
End With
Set rpdLineData = Nothing
Set objRapidPrim = Nothing
Set objRapidLine = Nothing
Set objGPData = Nothing

End Sub


Public Sub CreatePressBase()

Dim objRapidLine As clsRapidLine
Dim objGPData As clsGlassPlotData

Dim i As Long, bc As Long, ac As Long

Set objRapidLine = New clsRapidLine
Set objGPData = New clsGlassPlotData
With rpdDoc.DataBase
  bc = .Count
  objGPData.Calc
  objRapidLine.AddG gpdObject:=objGPData
  .Item(.Count - 1).Selected = True
  .RemoveSelect
  ac = .Count - 1
  For i = bc To ac
    With .Item(i)
      .Edit = True
      .LayerNo = lngPre(4)
      .PenStyle = zwStyleSDash
      .PenWidth = zwWidthThin
      .Edit = False
    End With
  Next
End With
Set objRapidLine = Nothing
Set objGPData = Nothing

End Sub


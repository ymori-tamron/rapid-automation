VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsBalsamPlotData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False


Option Explicit
Option Base 1
Option Compare Binary

Private dblOutMD As Double
Private dblOutBZ As Double, dblOutBH As Double
Private dblOutCX(2) As Double, dblOutCY(2) As Double
Private dblOutCZ(2) As Double, dblOutCH(2) As Double
Private dblOutPX(2) As Double, dblOutPY(2) As Double


Private Sub Class_Initialize()

Dim j As Long

dblOutMD = 0
dblOutBZ = 0: dblOutBH = 0
For j = 1 To 2
  dblOutCX(j) = 0: dblOutCY(j) = 0
  dblOutCZ(j) = 0: dblOutCH(j) = 0
  dblOutPX(j) = 0: dblOutPY(j) = 0
Next

End Sub


Public Property Get dblMaxDiameter() As Double

dblMaxDiameter = dblOutMD

End Property


Public Property Get dblBalsamZ() As Double

dblBalsamZ = dblOutBZ

End Property


Public Property Get dblBalsamH() As Double

dblBalsamH = dblOutBH

End Property


Public Property Get dblCenterX(ByVal lngRn As Long) As Double

dblCenterX = dblOutCX(lngRn)

End Property


Public Property Get dblCenterY(ByVal lngRn As Long) As Double

dblCenterY = dblOutCY(lngRn)

End Property


Public Property Get dblCornerZ(ByVal lngRn As Long) As Double

dblCornerZ = dblOutCZ(lngRn)

End Property


Public Property Get dblCornerH(ByVal lngRn As Long) As Double

dblCornerH = dblOutCH(lngRn)

End Property


Public Property Get dblPickupX(ByVal lngRn As Long) As Double

dblPickupX = dblOutPX(lngRn)

End Property


Public Property Get dblPickupY(ByVal lngRn As Long) As Double

dblPickupY = dblOutPY(lngRn)

End Property


Public Sub PreCalc()

Dim prtName(2) As String

prtName(1) = Cells(8, 3).Value
prtName(2) = Cells(9, 3).Value
If Left(ActiveSheet.Name, 1) = "B" Then
  If Left(Cells(8, 3).Value, 1) = "M" Then
    prtName(1) = "G" & Mid(Cells(8, 3).Value, 2, 2)
  End If
  If Left(Cells(9, 3).Value, 1) = "M" Then
    prtName(2) = "G" & Mid(Cells(9, 3).Value, 2, 2)
  End If
End If

If Worksheets(prtName(1)).Cells(8, 6).Value > _
   Worksheets(prtName(2)).Cells(8, 6).Value Then
  dblOutMD = Worksheets(prtName(1)).Cells(8, 6).Value
Else
  dblOutMD = Worksheets(prtName(2)).Cells(8, 6).Value
End If

End Sub


Public Sub CreateElement()

Dim rpdLineData As zwDrawCAD.LineData
Dim objRapidPrim As clsRapidPrim
Dim objRapidLine As clsRapidLine
Dim objGPData1 As clsGlassPlotData
Dim objGPData2 As clsGlassPlotData
Dim objGPData3 As clsGlassPlotData

Dim i As Long, j As Long
Dim bc As Long, ac As Long
Dim sft(3) As Double
Dim asName As String, prtName(3) As String

asName = ActiveSheet.Name
prtName(1) = Cells(8, 3).Value
prtName(2) = Cells(9, 3).Value
If Left(asName, 1) = "B" Then
  If Left(Cells(8, 3).Value, 1) = "M" Then
    prtName(1) = "G" & Mid(Cells(8, 3).Value, 2, 2)
  End If
  If Left(Cells(9, 3).Value, 1) = "M" Then
    prtName(2) = "G" & Mid(Cells(9, 3).Value, 2, 2)
  End If
End If
sft(1) = Worksheets(prtName(2)).Cells(10, 3).Value / 2
sft(2) = Worksheets(prtName(1)).Cells(10, 3).Value / 2

Set rpdLineData = New zwDrawCAD.LineData
Set objRapidPrim = New clsRapidPrim
Set objRapidLine = New clsRapidLine
Set objGPData1 = New clsGlassPlotData
Set objGPData2 = New clsGlassPlotData
Set objGPData3 = New clsGlassPlotData
With rpdDoc.DataBase
  bc = .Count
  Worksheets(prtName(1)).Activate
  objGPData1.Calc
  objRapidLine.AddG gpdObject:=objGPData1
  Worksheets(asName).Activate
  If Left(asName, 1) = "B" Or _
     Left(asName, 1) = "M" And _
     Left(prtName(1), 1) = "G" Then
    .Item(.Count - 1).Selected = True
    .RemoveSelect
  End If
  ac = .Count - 1
  For i = bc To ac
    .Item(i).Selected = True
  Next
  .MoveS dX:=-sft(1), dY:=0, KeepSelect:=zwKeepSelectNon
  
  bc = .Count
  Worksheets(prtName(2)).Activate
  objGPData2.Calc
  objRapidLine.AddG gpdObject:=objGPData2
  Worksheets(asName).Activate
  If Left(asName, 1) = "B" Or _
     Left(asName, 1) = "M" And _
     Left(prtName(2), 1) = "G" Then
    .Item(.Count - 1).Selected = True
    .RemoveSelect
  End If
  ac = .Count - 1
  For i = bc To ac
    .Item(i).Selected = True
  Next
  .MoveS dX:=sft(2), dY:=0, KeepSelect:=zwKeepSelectNon
End With

With objGPData1
  dblOutCX(1) = .dblCenterX(1) - sft(1)
  dblOutCY(1) = .dblCenterY(1)
  dblOutCZ(1) = .dblCornerZ(2, 1)
  dblOutCH(1) = .dblCornerH(2, 1)
  dblOutPX(1) = (.dblCornerZ(2, 1) + .dblCornerZ(2, 2)) / 2 + sft(2)
  dblOutPY(1) = .dblCornerH(2, 1) * 0.9
End With

With objGPData2
  dblOutCX(2) = .dblCenterX(2) + sft(2)
  dblOutCY(2) = .dblCenterY(2)
  dblOutCZ(2) = .dblCornerZ(2, 2)
  dblOutCH(2) = .dblCornerH(2, 2)
  dblOutPX(2) = (.dblCornerZ(2, 1) + .dblCornerZ(2, 2)) / 2 - sft(1)
  dblOutPY(2) = .dblCornerH(2, 2) * 0.9
End With

Select Case Left(asName, 1)
  Case "B"
    If dblOutCH(1) > dblOutCH(2) Then
      dblOutBH = dblOutCH(2) + 0.2
      With objGPData1
        If .dblCornerH(1, 2) - dblOutBH > 0 Then
          If .dblSurfaceH(2) <> 0 And _
             .dblSurfaceH(2) - dblOutBH < 0 Then
            dblOutBZ = .dblCornerZ(1, 2) - sft(1) * 2
          Else
            dblOutBZ = _
              Sag(Radius:=.dblRadius(2), Diameter:=dblOutBH * 2) - sft(1) * 2
          End If
        Else
          dblOutBH = 0
        End If
      End With
    Else
      dblOutBH = dblOutCH(1) + 0.2
      With objGPData2
        If .dblCornerH(1, 1) - dblOutBH > 0 Then
          If .dblSurfaceH(1) <> 0 And _
             .dblSurfaceH(1) - dblOutBH < 0 Then
            dblOutBZ = .dblCornerZ(1, 1) + sft(2) * 2
          Else
            dblOutBZ = _
              Sag(Radius:=.dblRadius(1), Diameter:=dblOutBH * 2) + sft(2) * 2
          End If
        Else
          dblOutBH = 0
        End If
      End With
    End If
    
    For j = 1 To 2
      If Left(Cells(j + 7, 3).Value, 1) = "M" Then
        prtName(3) = "H" & Mid(Cells(j + 7, 3).Value, 2, 2) & "(ASP)"
        sft(3) = Worksheets(prtName(3)).Cells(10, 3).Value / 2
        With rpdDoc.DataBase
          bc = .Count
          Worksheets(prtName(3)).Activate
          objGPData3.Calc
          objRapidLine.AddG gpdObject:=objGPData3
          Worksheets(asName).Activate
          ac = .Count - 1
          For i = bc To ac
            .Item(i).Selected = True
          Next
          .MoveS _
            dX:=(sft(1) + sft(2) + sft(3)) * (j * 2 - 3), _
            dY:=0, _
            KeepSelect:=zwKeepSelectNon
        End With
        
        dblOutCX(j) = dblOutCX(j) + (sft(3) * 2) * (j * 2 - 3)
        dblOutCZ(j) = dblOutCZ(j) - (sft(3) * 2) * (j * 2 - 3)
        dblOutPX(j) = dblOutPX(j) - (sft(3) * 2) * (j * 2 - 3)
        If (dblOutCH(1) - dblOutCH(2)) * (j * (-2) + 3) < 0 Then
          dblOutBZ = dblOutBZ - (sft(3) * 2) * (j * 2 - 3)
        End If
      Else
        objProgCtrl.AddBar lngDivCount:=1
      End If
    Next
    
  Case "M"
    If Left(prtName(1), 1) = "G" Then
      dblOutPY(2) = objGPData2.dblCornerH(2, 2) - 0.1
      
      With objGPData1
        If .dblSurfaceH(2) <> 0 Then
          dblOutBZ = .dblCornerZ(1, 2) - sft(1) * 2
          dblOutBH = .dblCornerH(1, 2)
        Else
          dblOutBH = 0
        End If
      End With
    ElseIf Left(prtName(2), 1) = "G" Then
      dblOutPY(1) = objGPData1.dblCornerH(2, 1) - 0.1
      
      With objGPData2
        If .dblSurfaceH(1) <> 0 Then
          dblOutBZ = .dblCornerZ(1, 1) + sft(2) * 2
          dblOutBH = .dblCornerH(1, 1)
        Else
          dblOutBH = 0
        End If
      End With
    End If
    
    objProgCtrl.AddBar lngDivCount:=1
    objProgCtrl.AddBar lngDivCount:=1
End Select

With rpdDoc
  .Views(0).Redraw
  rpdLineData.Start.x = dblOutCX(1) - 5 / .DrawInfo.Numerator
  rpdLineData.Start.y = dblOutCY(1)
  rpdLineData.End.x = dblOutCX(2) + 5 / .DrawInfo.Numerator
  rpdLineData.End.y = dblOutCY(2)
  objRapidPrim.SetUp prmObject:=rpdLineData, lngLayerNo:=CLng(Alpha(11)), _
                     lngPenStyle:=zwStyleS1DashDot, lngPenWidth:=zwWidthThin
  .DataBase.Add Prim:=rpdLineData
End With
Set rpdLineData = Nothing
Set objRapidPrim = Nothing
Set objRapidLine = Nothing
Set objGPData1 = Nothing
Set objGPData2 = Nothing
Set objGPData3 = Nothing

End Sub


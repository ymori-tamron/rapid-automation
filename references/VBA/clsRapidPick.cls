VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsRapidPick"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False


Option Explicit
Option Base 1
Option Compare Binary


Public Sub AddG(ByRef gpdObject As clsGlassPlotData)

Dim rpdPickData As zwDrawCAD.PickupData
Dim rpdTextData As zwDrawCAD.TextData
Dim rpdDPoint1 As zwDrawCAD.DPoint
Dim rpdDPoint2 As zwDrawCAD.DPoint
Dim rpdDPoint3 As zwDrawCAD.DPoint
Dim rpdDPoint4 As zwDrawCAD.DPoint
Dim objRapidPrim As clsRapidPrim
Dim objRapidParam As clsRapidParam
Dim objRapidTol As clsRapidTol

Dim i As Long, j As Long, bc As Long
Dim yy As Double, zz As Double
Dim ww As Double, tt As Double
Dim xx(3) As Double
Dim txt As String

Set rpdDPoint1 = New zwDrawCAD.DPoint
Set rpdDPoint2 = New zwDrawCAD.DPoint
Set rpdDPoint3 = New zwDrawCAD.DPoint
Set rpdDPoint4 = New zwDrawCAD.DPoint
Set objRapidPrim = New clsRapidPrim
Set objRapidParam = New clsRapidParam
objRapidParam.Import strConfig:="Param"
With gpdObject
  For i = 1 To objRapidParam.lngMaxCount
    Select Case objRapidParam.strItem(i)
      Case "Rê°ñ@"
        objProgCtrl.AddBar lngDivCount:=2
        xx(1) = objRapidParam.dblXY(i, 1) / rpdDoc.DrawInfo.Numerator
        yy = objRapidParam.dblXY(i, 2) / rpdDoc.DrawInfo.Numerator
        ww = objRapidParam.dblHW(i, 2) / rpdDoc.DrawInfo.Numerator
        bc = objRapidParam.lngVal(i, 2)
        
        For j = 1 To 2
          If Left(ActiveSheet.Name, 1) = "G" Or _
             Left(ActiveSheet.Name, 1) = "H" And _
             .blnAsp(j) Then
            txt = " R" & CStr(j) & "=" & Space(bc)
            xx(2) = xx(1) + (Len(txt) * 2) * (ww / 4)
            xx(3) = xx(1) + Len(txt) * (ww / 4)
            zz = objRapidParam.dblXY(i + j, 1) * (ww / 4)
            
            rpdDPoint1.x = .dblCenterX(j)
            rpdDPoint1.y = .dblCenterY(j)
            If .dblCornerZ(1, j) * (j * (-2) + 3) < 0 Then
              rpdDPoint2.x = .dblCenterX(j) + .dblCornerZ(1, j) + xx(1) * (j * 2 - 3)
              rpdDPoint3.x = .dblCenterX(j) + .dblCornerZ(1, j) + xx(2) * (j * 2 - 3)
              rpdDPoint4.x = .dblCenterX(j) + .dblCornerZ(1, j) + xx(3) * (j * 2 - 3)
            Else
              rpdDPoint2.x = .dblCenterX(j) + xx(1) * (j * 2 - 3)
              rpdDPoint3.x = .dblCenterX(j) + xx(2) * (j * 2 - 3)
              rpdDPoint4.x = .dblCenterX(j) + xx(3) * (j * 2 - 3)
            End If
            rpdDPoint2.y = .dblCenterY(j) - .dblCornerH(2, j) - yy
            rpdDPoint3.y = .dblCenterY(j) - .dblCornerH(2, j) - yy
            rpdDPoint4.y = .dblCenterY(j) - .dblCornerH(2, j) - yy
            
            objRapidParam.SetUp strParam:="PickParam"
            rpdDoc.PickupParam.Height = objRapidParam.dblHW(i, 1)
            rpdDoc.PickupParam.Width = objRapidParam.dblHW(i, 2)
            
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
            If .blnAsp(j) Then
              rpdTextData.Text = "ASP"
            Else
              rpdTextData.Text = Cells(objRapidParam.lngVal(i + j, 1), _
                                       objRapidParam.lngVal(i + j, 2)).Text
            End If
            rpdTextData.FaceName = objRapidParam.strItem(1)
            rpdTextData.Height = objRapidParam.dblHW(i + j, 1) / rpdDoc.DrawInfo.Numerator
            rpdTextData.Width = objRapidParam.dblHW(i + j, 2) / rpdDoc.DrawInfo.Numerator
            rpdTextData.BasePosition = 1
            rpdTextData.BasePoint.x = rpdDPoint4.x + zz
            rpdTextData.BasePoint.y = rpdDPoint4.y
            objRapidPrim.SetUp prmObject:=rpdTextData, lngLayerNo:=CLng(Alpha(12)), _
                               lngPenStyle:=zwStyleSolid, lngPenWidth:=zwWidthThin
            rpdDoc.DataBase.Add Prim:=rpdTextData
            Set rpdTextData = Nothing
          End If
        Next
        
      Case "Cê°ñ@"
        objProgCtrl.AddBar lngDivCount:=2
        xx(1) = objRapidParam.dblXY(i, 1) / rpdDoc.DrawInfo.Numerator
        yy = objRapidParam.dblXY(i, 2) / rpdDoc.DrawInfo.Numerator
        ww = objRapidParam.dblHW(i, 2) / rpdDoc.DrawInfo.Numerator
        bc = objRapidParam.lngVal(i, 2)
        
        For j = 1 To 2
          If .dblCornerH(2, j) - .dblCornerH(1, j) > 0 And _
             Left(ActiveSheet.Name, 1) = "G" Then
            Set objRapidTol = New clsRapidTol
            txt = Format(.dblCornerH(2, j) - .dblCornerH(1, j), """C""0.0#")
            xx(2) = xx(1) + (Len(txt) * 2) * (ww / 4)
            xx(3) = xx(1) + Len(txt) * (ww / 4)
            zz = 0: tt = 0
            objRapidTol.Check rpObject:=objRapidParam, lngRow:=i + j
            If objRapidTol.strKeyWord = "Num" Then
              zz = bc * (ww / 4)
              tt = (Len(txt) - bc) * (ww / 4)
              txt = txt & Space(bc)
            ElseIf objRapidTol.strKeyWord <> "Empty" Then
              zz = LenB(StrConv(objRapidTol.strKeyWord, vbFromUnicode)) * (ww / 4)
              txt = txt & objRapidTol.strKeyWord
            End If
            
            rpdDPoint1.x = .dblCenterX(j) + (.dblCornerZ(1, j) + .dblCornerZ(2, j)) / 2
            rpdDPoint1.y = .dblCenterY(j) + (.dblCornerH(1, j) + .dblCornerH(2, j)) / 2
            rpdDPoint2.x = .dblCenterX(j) + .dblCornerZ(1, j) + xx(1) * (j * 2 - 3)
            rpdDPoint2.y = .dblCenterY(j) + .dblCornerH(2, j) + yy
            rpdDPoint3.x = .dblCenterX(j) + .dblCornerZ(1, j) + (xx(2) + zz * 2) * (j * 2 - 3)
            rpdDPoint3.y = .dblCenterY(j) + .dblCornerH(2, j) + yy
            rpdDPoint4.x = .dblCenterX(j) + .dblCornerZ(1, j) + (xx(3) + zz) * (j * 2 - 3)
            rpdDPoint4.y = .dblCenterY(j) + .dblCornerH(2, j) + yy
            
            objRapidParam.SetUp strParam:="PickParam"
            rpdDoc.PickupParam.Height = objRapidParam.dblHW(i, 1)
            rpdDoc.PickupParam.Width = objRapidParam.dblHW(i, 2)
            
            Set rpdPickData = New zwDrawCAD.PickupData
            rpdPickData.CreateEx doc:=rpdDoc, _
                                 StartPoint:=rpdDPoint1, CornerPoint:=rpdDPoint2, _
                                 EndPoint:=rpdDPoint3, TextPoint:=rpdDPoint4, _
                                 sText:=txt, Arrange:=zwPickupTextFreeCorrect
            objRapidPrim.SetUp prmObject:=rpdPickData, lngLayerNo:=CLng(Alpha(13)), _
                               lngPenStyle:=zwStyleSolid, lngPenWidth:=zwWidthThin
            rpdDoc.DataBase.Add Prim:=rpdPickData
            objRapidTol.AddG rpObject:=objRapidParam, lngRow:=i + j, _
                             dblBaseX:=rpdDPoint4.x, dblBaseY:=rpdDPoint4.y, _
                             dblShift:=tt, dblAngle:=0
            Set rpdPickData = Nothing
            Set objRapidTol = Nothing
          End If
          
          If .dblSurfaceH(j) <> 0 And _
             InStr(ActiveSheet.Name, "ASP") = 0 And _
             Left(ActiveSheet.Name, 1) = "G" Then
            txt = "äOåaÇ∆íºäpñ éÊÇÕìØêcâ¡çHÇÃéñ"
            xx(2) = xx(1) + (LenB(txt) * 2) * ((ww / 1.5) / 4)
            xx(3) = xx(1) + LenB(txt) * ((ww / 1.5) / 4)
            
            rpdDPoint1.x = .dblCenterX(j) + .dblCornerZ(1, j)
            rpdDPoint1.y = .dblCenterY(j) + (.dblSurfaceH(j) + .dblCornerH(1, j)) / 2
            rpdDPoint2.x = .dblCenterX(j) + .dblCornerZ(1, j) + xx(1) * (j * 2 - 3)
            rpdDPoint2.y = .dblCenterY(j) + .dblCornerH(2, j) + yy / 2
            rpdDPoint3.x = .dblCenterX(j) + .dblCornerZ(1, j) + xx(2) * (j * 2 - 3)
            rpdDPoint3.y = .dblCenterY(j) + .dblCornerH(2, j) + yy / 2
            rpdDPoint4.x = .dblCenterX(j) + .dblCornerZ(1, j) + xx(3) * (j * 2 - 3)
            rpdDPoint4.y = .dblCenterY(j) + .dblCornerH(2, j) + yy / 2
            
            objRapidParam.SetUp strParam:="PickParam"
            rpdDoc.PickupParam.Height = objRapidParam.dblHW(i, 1) / 1.5
            rpdDoc.PickupParam.Width = objRapidParam.dblHW(i, 2) / 1.5
            
            Set rpdPickData = New zwDrawCAD.PickupData
            rpdPickData.CreateEx doc:=rpdDoc, _
                                 StartPoint:=rpdDPoint1, CornerPoint:=rpdDPoint2, _
                                 EndPoint:=rpdDPoint3, TextPoint:=rpdDPoint4, _
                                 sText:=txt, Arrange:=zwPickupTextFreeCorrect
            objRapidPrim.SetUp prmObject:=rpdPickData, lngLayerNo:=CLng(Alpha(13)), _
                               lngPenStyle:=zwStyleSolid, lngPenWidth:=zwWidthThin
            rpdDoc.DataBase.Add Prim:=rpdPickData
            Set rpdPickData = Nothing
          End If
        Next
    End Select
  Next
  objRapidParam.SetUp strParam:="PickParam"
End With
Set rpdDPoint1 = Nothing
Set rpdDPoint2 = Nothing
Set rpdDPoint3 = Nothing
Set rpdDPoint4 = Nothing
Set objRapidPrim = Nothing
Set objRapidParam = Nothing

End Sub


Public Sub AddP(ByRef ppdObject As clsPressPlotData)

Dim rpdPickData As zwDrawCAD.PickupData
Dim rpdLineData As zwDrawCAD.LineData
Dim rpdTextData As zwDrawCAD.TextData
Dim rpdDPoint1 As zwDrawCAD.DPoint
Dim rpdDPoint2 As zwDrawCAD.DPoint
Dim rpdDPoint3 As zwDrawCAD.DPoint
Dim rpdDPoint4 As zwDrawCAD.DPoint
Dim objRapidPrim As clsRapidPrim
Dim objRapidParam As clsRapidParam

Dim i As Long, j As Long, bc As Long
Dim yy As Double, zz As Double, ww As Double
Dim xx(3) As Double
Dim txt As String

Set rpdDPoint1 = New zwDrawCAD.DPoint
Set rpdDPoint2 = New zwDrawCAD.DPoint
Set rpdDPoint3 = New zwDrawCAD.DPoint
Set rpdDPoint4 = New zwDrawCAD.DPoint
Set objRapidPrim = New clsRapidPrim
Set objRapidParam = New clsRapidParam
objRapidParam.Import strConfig:="Param"
With ppdObject
  For i = 1 To objRapidParam.lngMaxCount
    Select Case objRapidParam.strItem(i)
      Case "Rê°ñ@"
        objProgCtrl.AddBar lngDivCount:=2
        xx(1) = objRapidParam.dblXY(i, 1) / rpdDoc.DrawInfo.Numerator
        yy = objRapidParam.dblXY(i, 2) / rpdDoc.DrawInfo.Numerator
        ww = objRapidParam.dblHW(i, 2) / rpdDoc.DrawInfo.Numerator
        bc = objRapidParam.lngVal(i, 2)
        
        For j = 1 To 2
          txt = " R" & CStr(j) & "=" & Space(bc)
          xx(2) = xx(1) + (Len(txt) * 2) * (ww / 4)
          xx(3) = xx(1) + Len(txt) * (ww / 4)
          zz = objRapidParam.dblXY(i + j, 1) * (ww / 4)
          
          If .dblCornerZ(1, j) * (j * (-2) + 3) < 0 Then
            rpdDPoint2.x = .dblCenterX(j) + .dblCornerZ(1, j) + xx(1) * (j * 2 - 3)
            rpdDPoint3.x = .dblCenterX(j) + .dblCornerZ(1, j) + xx(2) * (j * 2 - 3)
            rpdDPoint4.x = .dblCenterX(j) + .dblCornerZ(1, j) + xx(3) * (j * 2 - 3)
          Else
            rpdDPoint2.x = .dblCenterX(j) + xx(1) * (j * 2 - 3)
            rpdDPoint3.x = .dblCenterX(j) + xx(2) * (j * 2 - 3)
            rpdDPoint4.x = .dblCenterX(j) + xx(3) * (j * 2 - 3)
          End If
          rpdDPoint2.y = .dblCenterY(j) - .dblCornerH(2, j) - yy
          rpdDPoint3.y = .dblCenterY(j) - .dblCornerH(2, j) - yy
          rpdDPoint4.y = .dblCenterY(j) - .dblCornerH(2, j) - yy
          
          Set rpdLineData = New zwDrawCAD.LineData
          rpdLineData.Start.x = rpdDPoint2.x
          rpdLineData.Start.y = rpdDPoint2.y
          rpdLineData.End.x = rpdDPoint3.x
          rpdLineData.End.y = rpdDPoint3.y
          objRapidPrim.SetUp prmObject:=rpdLineData, lngLayerNo:=lngPre(4), _
                             lngPenStyle:=zwStyleSolid, lngPenWidth:=zwWidthThin
          rpdDoc.DataBase.Add Prim:=rpdLineData
          Set rpdLineData = Nothing
          
          Set rpdTextData = New zwDrawCAD.TextData
          rpdTextData.Text = txt
          rpdTextData.FaceName = objRapidParam.strItem(1)
          rpdTextData.Height = objRapidParam.dblHW(i, 1) / rpdDoc.DrawInfo.Numerator
          rpdTextData.Width = objRapidParam.dblHW(i, 2) / rpdDoc.DrawInfo.Numerator
          rpdTextData.BasePosition = 1
          rpdTextData.BasePoint.x = rpdDPoint4.x
          rpdTextData.BasePoint.y = rpdDPoint4.y
          objRapidPrim.SetUp prmObject:=rpdTextData, lngLayerNo:=lngPre(4), _
                             lngPenStyle:=zwStyleSolid, lngPenWidth:=zwWidthThin
          rpdDoc.DataBase.Add Prim:=rpdTextData
          Set rpdTextData = Nothing
          
          Set rpdTextData = New zwDrawCAD.TextData
          If Abs(.dblRadius(j)) > 1000000000# Then
            rpdTextData.Text = "Inf"
          Else
            rpdTextData.Text = Format(.dblRadius(j) * (j * (-2) + 3), "+0.0;-0.0")
          End If
          rpdTextData.FaceName = objRapidParam.strItem(1)
          rpdTextData.Height = objRapidParam.dblHW(i + j, 1) / rpdDoc.DrawInfo.Numerator
          rpdTextData.Width = objRapidParam.dblHW(i + j, 2) / rpdDoc.DrawInfo.Numerator
          rpdTextData.BasePosition = 1
          rpdTextData.BasePoint.x = rpdDPoint4.x + zz
          rpdTextData.BasePoint.y = rpdDPoint4.y
          objRapidPrim.SetUp prmObject:=rpdTextData, lngLayerNo:=lngPre(3), _
                             lngPenStyle:=zwStyleSolid, lngPenWidth:=zwWidthThin
          rpdDoc.DataBase.Add Prim:=rpdTextData
          Set rpdTextData = Nothing
        Next
        
      Case "Cê°ñ@"
        objProgCtrl.AddBar lngDivCount:=2
        xx(1) = objRapidParam.dblXY(i, 1) / rpdDoc.DrawInfo.Numerator
        yy = objRapidParam.dblXY(i, 2) / rpdDoc.DrawInfo.Numerator
        ww = objRapidParam.dblHW(i, 2) / rpdDoc.DrawInfo.Numerator
        
        For j = 1 To 2
          If .dblCornerH(2, j) - .dblCornerH(1, j) > 0 Then
            txt = Format(.dblCornerH(2, j) - .dblCornerH(1, j), """C""0.0#")
            xx(2) = xx(1) + (Len(txt) * 2) * (ww / 4)
            xx(3) = xx(1) + Len(txt) * (ww / 4)
            
            rpdDPoint1.x = .dblCenterX(j) + (.dblCornerZ(1, j) + .dblCornerZ(2, j)) / 2
            rpdDPoint1.y = .dblCenterY(j) + (.dblCornerH(1, j) + .dblCornerH(2, j)) / 2
            rpdDPoint2.x = .dblCenterX(j) + .dblCornerZ(1, j) + xx(1) * (j * 2 - 3)
            rpdDPoint2.y = .dblCenterY(j) + .dblCornerH(2, j) + yy
            rpdDPoint3.x = .dblCenterX(j) + .dblCornerZ(1, j) + xx(2) * (j * 2 - 3)
            rpdDPoint3.y = .dblCenterY(j) + .dblCornerH(2, j) + yy
            rpdDPoint4.x = .dblCenterX(j) + .dblCornerZ(1, j) + xx(3) * (j * 2 - 3)
            rpdDPoint4.y = .dblCenterY(j) + .dblCornerH(2, j) + yy
            
            objRapidParam.SetUp strParam:="PickParam"
            rpdDoc.PickupParam.Height = objRapidParam.dblHW(i, 1)
            rpdDoc.PickupParam.Width = objRapidParam.dblHW(i, 2)
            
            Set rpdPickData = New zwDrawCAD.PickupData
            rpdPickData.CreateEx doc:=rpdDoc, _
                                 StartPoint:=rpdDPoint1, CornerPoint:=rpdDPoint2, _
                                 EndPoint:=rpdDPoint3, TextPoint:=rpdDPoint4, _
                                 sText:=txt, Arrange:=zwPickupTextFreeCorrect
            objRapidPrim.SetUp prmObject:=rpdPickData, lngLayerNo:=lngPre(4), _
                               lngPenStyle:=zwStyleSolid, lngPenWidth:=zwWidthThin
            rpdDoc.DataBase.Add Prim:=rpdPickData
            Set rpdPickData = Nothing
          End If
        Next
    End Select
  Next
  objRapidParam.SetUp strParam:="PickParam"
End With
Set rpdDPoint1 = Nothing
Set rpdDPoint2 = Nothing
Set rpdDPoint3 = Nothing
Set rpdDPoint4 = Nothing
Set objRapidPrim = Nothing
Set objRapidParam = Nothing

End Sub


Public Sub AddB(ByRef bpdObject As clsBalsamPlotData)

Dim rpdBalloonData As zwDrawCAD.BalloonData
Dim rpdTextData As zwDrawCAD.TextData
Dim rpdDPoint1 As zwDrawCAD.DPoint
Dim rpdDPoint2 As zwDrawCAD.DPoint
Dim objRapidPrim As clsRapidPrim
Dim objRapidParam As clsRapidParam

Dim i As Long, j As Long
Dim xx As Double, yy As Double
Dim hh As Double, ww As Double, ff As Double
Dim prtName(3) As String, txt(2) As String

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

Set rpdDPoint1 = New zwDrawCAD.DPoint
Set rpdDPoint2 = New zwDrawCAD.DPoint
Set objRapidPrim = New clsRapidPrim
Set objRapidParam = New clsRapidParam
objRapidParam.Import strConfig:="Param"
With bpdObject
  For i = 1 To objRapidParam.lngMaxCount
    Select Case objRapidParam.strItem(i)
      Case "Rê°ñ@"
        objProgCtrl.AddBar lngDivCount:=1
        xx = objRapidParam.dblXY(i, 1) / rpdDoc.DrawInfo.Numerator
        yy = objRapidParam.dblXY(i, 2) / rpdDoc.DrawInfo.Numerator
        hh = objRapidParam.dblHW(i, 1) / rpdDoc.DrawInfo.Numerator
        ww = objRapidParam.dblHW(i, 2) / rpdDoc.DrawInfo.Numerator
        
        For j = 1 To 2
          Select Case Left(ActiveSheet.Name, 1)
            Case "B"
              ff = 1
              txt(1) = ""
              txt(2) = Mid(Trim(Left(prtName(j), 3)) & Space(4), 1, 4)
            Case "M"
              ff = 0.75
              prtName(3) = Worksheets(prtName(j)).Cells(3, 3).Value
              txt(1) = Mid(Trim(Left(prtName(3), 3)) & Space(4), 1, 4)
              txt(2) = Trim(Mid(prtName(3), 4))
          End Select
          
          rpdDPoint1.x = .dblCenterX(j) + .dblPickupX(j)
          rpdDPoint1.y = .dblCenterY(j) - .dblPickupY(j)
          If .dblCornerZ(j) * (j * (-2) + 3) < 0 Then
            rpdDPoint2.x = .dblCenterX(j) + .dblCornerZ(j) + (xx + ww) * (j * 2 - 3)
          Else
            rpdDPoint2.x = .dblCenterX(j) + (xx + ww) * (j * 2 - 3)
          End If
          rpdDPoint2.y = .dblCenterY(j) - .dblMaxDiameter / 2 - yy
          
          objRapidParam.SetUp strParam:="PickParam"
          rpdDoc.PickupParam.Height = objRapidParam.dblHW(i, 1)
          rpdDoc.PickupParam.Width = objRapidParam.dblHW(i, 2)
          rpdDoc.PickupParam.MarkType = zwMarkBlackCircle
          rpdDoc.PickupParam.BalloonType = zwBalloonCircle
          
          Set rpdBalloonData = New zwDrawCAD.BalloonData
          rpdBalloonData.Create doc:=rpdDoc, _
                                StartPoint:=rpdDPoint1, CenterPoint:=rpdDPoint2, _
                                dRadius:=hh * (Sqr(2) + 0.1), sText:=" "
          objRapidPrim.SetUp prmObject:=rpdBalloonData, lngLayerNo:=CLng(Alpha(13)), _
                             lngPenStyle:=zwStyleSolid, lngPenWidth:=zwWidthThin
          rpdDoc.DataBase.Add Prim:=rpdBalloonData
          Set rpdBalloonData = Nothing
          
          If txt(1) <> "" Then
            rpdDPoint2.y = rpdDPoint2.y - hh * (ff * 2 - 1) / 2
            
            Set rpdTextData = New zwDrawCAD.TextData
            rpdTextData.Text = txt(1)
            rpdTextData.FaceName = objRapidParam.strItem(1)
            rpdTextData.Height = hh * ff
            rpdTextData.Width = ww * ff
            rpdTextData.BasePosition = 1
            rpdTextData.BasePoint.x = rpdDPoint2.x
            rpdTextData.BasePoint.y = rpdDPoint2.y + hh * ff
            objRapidPrim.SetUp prmObject:=rpdTextData, lngLayerNo:=CLng(Alpha(13)), _
                               lngPenStyle:=zwStyleSolid, lngPenWidth:=zwWidthThin
            rpdDoc.DataBase.Add Prim:=rpdTextData
          Set rpdTextData = Nothing
          End If
          
          Set rpdTextData = New zwDrawCAD.TextData
          rpdTextData.Text = txt(2)
          rpdTextData.FaceName = objRapidParam.strItem(1)
          rpdTextData.Height = hh * ff
          rpdTextData.Width = ww * ff
          rpdTextData.BasePosition = 1
          rpdTextData.BasePoint.x = rpdDPoint2.x
          rpdTextData.BasePoint.y = rpdDPoint2.y
          objRapidPrim.SetUp prmObject:=rpdTextData, lngLayerNo:=CLng(Alpha(13)), _
                             lngPenStyle:=zwStyleSolid, lngPenWidth:=zwWidthThin
          rpdDoc.DataBase.Add Prim:=rpdTextData
          Set rpdTextData = Nothing
          
          Set rpdTextData = New zwDrawCAD.TextData
          If Not CBool(fCheck(5)) Then
            rpdTextData.Text = "- 00"
          Else
            rpdTextData.Text = "-   "
          End If
          rpdTextData.FaceName = objRapidParam.strItem(1)
          rpdTextData.Height = hh * ff
          rpdTextData.Width = ww * ff
          rpdTextData.BasePosition = 7
          rpdTextData.BasePoint.x = rpdDPoint2.x
          rpdTextData.BasePoint.y = rpdDPoint2.y
          objRapidPrim.SetUp prmObject:=rpdTextData, lngLayerNo:=CLng(Alpha(13)), _
                             lngPenStyle:=zwStyleSolid, lngPenWidth:=zwWidthThin
          rpdDoc.DataBase.Add Prim:=rpdTextData
          Set rpdTextData = Nothing
        Next
    End Select
  Next
  objRapidParam.SetUp strParam:="PickParam"
End With
Set rpdDPoint1 = Nothing
Set rpdDPoint2 = Nothing
Set objRapidPrim = Nothing
Set objRapidParam = Nothing

End Sub


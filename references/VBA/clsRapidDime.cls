VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsRapidDime"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False


Option Explicit
Option Base 1
Option Compare Binary


Public Sub AddG(ByRef gpdObject As clsGlassPlotData)

Dim rpdDimeData As zwDrawCAD.DimensionData
Dim rpdDPoint1 As zwDrawCAD.DPoint
Dim rpdDPoint2 As zwDrawCAD.DPoint
Dim rpdDPoint3 As zwDrawCAD.DPoint
Dim rpdDPoint4 As zwDrawCAD.DPoint
Dim objRapidPrim As clsRapidPrim
Dim objRapidParam As clsRapidParam
Dim objRapidTol As clsRapidTol

Dim i As Long, j As Long
Dim dd As Long, bc As Long
Dim aa As Double, xx As Double, yy As Double
Dim hh As Double, ww As Double, tt As Double

Set rpdDPoint1 = New zwDrawCAD.DPoint
Set rpdDPoint2 = New zwDrawCAD.DPoint
Set rpdDPoint3 = New zwDrawCAD.DPoint
Set rpdDPoint4 = New zwDrawCAD.DPoint
Set objRapidPrim = New clsRapidPrim
Set objRapidParam = New clsRapidParam
objRapidParam.Import strConfig:="Param"
aa = objRapidParam.dblXY(1, 1) / rpdDoc.DrawInfo.Numerator
With gpdObject
  For i = 1 To objRapidParam.lngMaxCount
    Select Case objRapidParam.strItem(i)
      Case "肉厚寸法"
        objProgCtrl.AddBar lngDivCount:=4
        yy = objRapidParam.dblXY(i, 2) / rpdDoc.DrawInfo.Numerator
        ww = objRapidParam.dblHW(i, 2) / rpdDoc.DrawInfo.Numerator
        bc = objRapidParam.lngVal(i, 2)
        
        rpdDPoint1.x = .dblCenterX(1)
        rpdDPoint1.y = .dblCenterY(1)
        rpdDPoint2.x = .dblCenterX(2)
        rpdDPoint2.y = .dblCenterY(2)
        rpdDPoint3.x = (.dblCenterX(1) + .dblCenterX(2)) / 2
        rpdDPoint3.y = .dblCenterY(1) - .dblCornerH(2, 1) - yy
        
        objRapidParam.SetUp strParam:="DimeParam"
        rpdDoc.DimensionParam.Height = objRapidParam.dblHW(i, 1)
        rpdDoc.DimensionParam.Width = objRapidParam.dblHW(i, 2)
        rpdDoc.DimensionParam.DecimalDigit = objRapidParam.lngVal(i, 1)
        
        Set rpdDimeData = New zwDrawCAD.DimensionData
        Set objRapidTol = New clsRapidTol
        rpdDimeData.CreateSingle doc:=rpdDoc, _
                                 PickStart:=rpdDPoint1, PickEnd:=rpdDPoint2, _
                                 PassPoint:=rpdDPoint3, dAngle:=0
        
        xx = Len(rpdDimeData.Text(0)) * (ww / 4) + aa + ww / 8
        tt = 0
        objRapidTol.Check rpObject:=objRapidParam, lngRow:=i + 1
        If objRapidTol.strKeyWord = "Num" Then
          xx = xx + bc * (ww / 4)
          tt = (Len(rpdDimeData.Text(0)) - bc) * (ww / 4)
          rpdDimeData.Text(0) = rpdDimeData.Text(0) & Space(bc)
        End If
        If (.dblCenterX(2) - .dblCenterX(1)) / 2 < xx Then
          rpdDPoint3.x = .dblCenterX(2) + xx
        End If
        rpdDPoint3.y = rpdDPoint3.y + 0.001 / rpdDoc.DrawInfo.Numerator
        rpdDimeData.SetTextPoint doc:=rpdDoc, lIndex:=0, _
                                 PutPoint:=rpdDPoint3, Arrange:=zwDimTextParallel
        
        objRapidPrim.SetUp prmObject:=rpdDimeData, lngLayerNo:=CLng(Alpha(13)), _
                           lngPenStyle:=zwStyleSolid, lngPenWidth:=zwWidthThin
        rpdDoc.DataBase.Add Prim:=rpdDimeData
        objRapidTol.AddG rpObject:=objRapidParam, lngRow:=i + 1, _
                         dblBaseX:=rpdDPoint3.x, dblBaseY:=rpdDPoint3.y, _
                         dblShift:=tt, dblAngle:=0
        Set rpdDimeData = Nothing
        Set objRapidTol = Nothing
        
      Case "⊿Ｈ寸法"
        objProgCtrl.AddBar lngDivCount:=4
        yy = objRapidParam.dblXY(i, 2) / rpdDoc.DrawInfo.Numerator
        ww = objRapidParam.dblHW(i, 2) / rpdDoc.DrawInfo.Numerator
        dd = objRapidParam.lngVal(i, 1)
        bc = objRapidParam.lngVal(i, 2)
        
        For j = 1 To 2
          If .dblSurfaceH(j) <> 0 And _
             Left(ActiveSheet.Name, 1) = "G" And _
             Cells(j + 10, 6).Value > 0 Then
            rpdDPoint1.x = .dblCenterX(j)
            rpdDPoint1.y = .dblCenterY(j)
            rpdDPoint2.x = .dblCenterX(j) + .dblCornerZ(1, j)
            rpdDPoint2.y = .dblCenterY(j) - .dblCornerH(1, j)
            rpdDPoint3.x = .dblCenterX(j) + .dblCornerZ(1, j) / 2
            rpdDPoint3.y = .dblCenterY(j) - .dblCornerH(2, j) - yy
            
            objRapidParam.SetUp strParam:="DimeParam"
            rpdDoc.DimensionParam.Height = objRapidParam.dblHW(i, 1)
            rpdDoc.DimensionParam.Width = objRapidParam.dblHW(i, 2)
            rpdDoc.DimensionParam.DecimalDigit = 0
            If .dblCornerZ(1, j) * (j * (-2) + 3) < 0 Then
              rpdDoc.DimensionParam.StartMarkType = zwMarkBlackCircle
              rpdDoc.DimensionParam.StartDistance = 0.001 / rpdDoc.DrawInfo.Numerator
            Else
              rpdDoc.DimensionParam.EndMarkType = zwMarkBlackCircle
              rpdDoc.DimensionParam.EndDistance = 0.001 / rpdDoc.DrawInfo.Numerator
            End If
            
            Set rpdDimeData = New zwDrawCAD.DimensionData
            Set objRapidTol = New clsRapidTol
            rpdDimeData.CreateSingle doc:=rpdDoc, _
                                     PickStart:=rpdDPoint1, PickEnd:=rpdDPoint2, _
                                     PassPoint:=rpdDPoint3, dAngle:=0
            
            rpdDimeData.Text(0) = Format(Abs(.dblCornerZ(1, j)), _
                                         "0." & String(dd, "0"))
            xx = Len(rpdDimeData.Text(0)) * (ww / 4) + aa + ww / 8
            tt = 0
            objRapidTol.Check rpObject:=objRapidParam, lngRow:=i + j
            If objRapidTol.strKeyWord = "Num" Then
              xx = xx + bc * (ww / 4)
              tt = (Len(rpdDimeData.Text(0)) - bc) * (ww / 4)
              rpdDimeData.Text(0) = rpdDimeData.Text(0) & Space(bc)
            End If
            If .dblCornerZ(1, j) * (j * (-2) + 3) < 0 Then
              rpdDPoint3.x = .dblCenterX(j) + .dblCornerZ(1, j) + xx * (j * 2 - 3)
            Else
              rpdDPoint3.x = .dblCenterX(j) + xx * (j * 2 - 3)
            End If
            rpdDPoint3.y = rpdDPoint3.y + 0.001 / rpdDoc.DrawInfo.Numerator
            rpdDimeData.SetTextPoint doc:=rpdDoc, lIndex:=0, _
                                     PutPoint:=rpdDPoint3, Arrange:=zwDimTextParallel
            
            objRapidPrim.SetUp prmObject:=rpdDimeData, lngLayerNo:=CLng(Alpha(13)), _
                               lngPenStyle:=zwStyleSolid, lngPenWidth:=zwWidthThin
            rpdDoc.DataBase.Add Prim:=rpdDimeData
            objRapidTol.AddG rpObject:=objRapidParam, lngRow:=i + j, _
                             dblBaseX:=rpdDPoint3.x, dblBaseY:=rpdDPoint3.y, _
                             dblShift:=tt, dblAngle:=0
            Set rpdDimeData = Nothing
            Set objRapidTol = Nothing
          End If
        Next
        
      Case "外径寸法"
        objProgCtrl.AddBar lngDivCount:=4
        xx = objRapidParam.dblXY(i, 1) / rpdDoc.DrawInfo.Numerator
        ww = objRapidParam.dblHW(i, 2) / rpdDoc.DrawInfo.Numerator
        bc = objRapidParam.lngVal(i, 2)
        
        If Left(ActiveSheet.Name, 1) = "G" Then
          rpdDPoint1.x = .dblCenterX(1) + .dblCornerZ(2, 1)
          rpdDPoint1.y = .dblCenterY(1) + .dblCornerH(2, 1)
          rpdDPoint2.x = .dblCenterX(1) + .dblCornerZ(2, 1)
          rpdDPoint2.y = .dblCenterY(1) - .dblCornerH(2, 1)
          If .dblCornerZ(1, 1) < 0 Then
            rpdDPoint3.x = .dblCenterX(1) + .dblCornerZ(1, 1) - xx
          Else
            rpdDPoint3.x = .dblCenterX(1) - xx
          End If
          rpdDPoint3.y = .dblCenterY(1)
          
          objRapidParam.SetUp strParam:="DimeParam"
          rpdDoc.DimensionParam.Height = objRapidParam.dblHW(i, 1)
          rpdDoc.DimensionParam.Width = objRapidParam.dblHW(i, 2)
          rpdDoc.DimensionParam.DecimalDigit = objRapidParam.lngVal(i, 1)
          
          Set rpdDimeData = New zwDrawCAD.DimensionData
          Set objRapidTol = New clsRapidTol
          rpdDimeData.CreateSingle doc:=rpdDoc, _
                                   PickStart:=rpdDPoint1, PickEnd:=rpdDPoint2, _
                                   PassPoint:=rpdDPoint3, dAngle:=90
          
          rpdDimeData.Text(0) = "φ" & rpdDimeData.Text(0)
          tt = 0
          objRapidTol.Check rpObject:=objRapidParam, lngRow:=i + 1
          If objRapidTol.strKeyWord = "Num" Then
            tt = (Len(rpdDimeData.Text(0)) + 1 - bc) * (ww / 4)
            rpdDimeData.Text(0) = rpdDimeData.Text(0) & Space(bc)
          ElseIf objRapidTol.strKeyWord <> "Empty" Then
            rpdDimeData.Text(0) = rpdDimeData.Text(0) & objRapidTol.strKeyWord
          End If
          
          objRapidPrim.SetUp prmObject:=rpdDimeData, lngLayerNo:=CLng(Alpha(13)), _
                             lngPenStyle:=zwStyleSolid, lngPenWidth:=zwWidthThin
          rpdDoc.DataBase.Add Prim:=rpdDimeData
          objRapidTol.AddG rpObject:=objRapidParam, lngRow:=i + 1, _
                           dblBaseX:=rpdDPoint3.x, dblBaseY:=rpdDPoint3.y, _
                           dblShift:=tt, dblAngle:=90
          Set rpdDimeData = Nothing
          Set objRapidTol = Nothing
        End If
        
      Case "内径寸法"
        objProgCtrl.AddBar lngDivCount:=4
        xx = objRapidParam.dblXY(i, 1) / rpdDoc.DrawInfo.Numerator
        hh = objRapidParam.dblHW(i, 1) / rpdDoc.DrawInfo.Numerator
        ww = objRapidParam.dblHW(i, 2) / rpdDoc.DrawInfo.Numerator
        bc = objRapidParam.lngVal(i, 2)
        
        For j = 1 To 2
          If .dblSurfaceH(j) <> 0 Then
            If Left(ActiveSheet.Name, 1) = "G" Or _
               Left(ActiveSheet.Name, 1) = "H" And _
               .blnAsp(j) Then
              rpdDPoint1.x = .dblCenterX(j) + .dblCornerZ(1, j)
              rpdDPoint1.y = .dblCenterY(j) + .dblSurfaceH(j)
              rpdDPoint2.x = .dblCenterX(j) + .dblCornerZ(1, j)
              rpdDPoint2.y = .dblCenterY(j) - .dblSurfaceH(j)
              If .dblCornerZ(1, j) * (j * (-2) + 3) < 0 Then
                rpdDPoint3.x = .dblCenterX(j) + .dblCornerZ(1, j) + xx * (j * 2 - 3)
              Else
                rpdDPoint3.x = .dblCenterX(j) + xx * (j * 2 - 3)
              End If
              rpdDPoint3.y = .dblCenterY(j)
              
              objRapidParam.SetUp strParam:="DimeParam"
              rpdDoc.DimensionParam.Height = objRapidParam.dblHW(i, 1)
              rpdDoc.DimensionParam.Width = objRapidParam.dblHW(i, 2)
              rpdDoc.DimensionParam.DecimalDigit = objRapidParam.lngVal(i, 1)
              
              Set rpdDimeData = New zwDrawCAD.DimensionData
              Set objRapidTol = New clsRapidTol
              rpdDimeData.CreateSingle doc:=rpdDoc, _
                                       PickStart:=rpdDPoint1, PickEnd:=rpdDPoint2, _
                                       PassPoint:=rpdDPoint3, dAngle:=90
              
              rpdDimeData.Text(0) = "φ" & rpdDimeData.Text(0)
              tt = 0
              objRapidTol.Check rpObject:=objRapidParam, lngRow:=i + j
              If objRapidTol.strKeyWord = "Num" Then
                tt = (Len(rpdDimeData.Text(0)) + 1 - bc) * (ww / 4)
                rpdDimeData.Text(0) = rpdDimeData.Text(0) & Space(bc)
              ElseIf objRapidTol.strKeyWord = "()" Then
                rpdDimeData.Text(0) = "(" & rpdDimeData.Text(0) & ")"
              ElseIf objRapidTol.strKeyWord <> "Empty" Then
                rpdDimeData.Text(0) = rpdDimeData.Text(0) & objRapidTol.strKeyWord
              End If
              
              objRapidPrim.SetUp prmObject:=rpdDimeData, lngLayerNo:=CLng(Alpha(13)), _
                                 lngPenStyle:=zwStyleSolid, lngPenWidth:=zwWidthThin
              rpdDoc.DataBase.Add Prim:=rpdDimeData
              objRapidTol.AddG rpObject:=objRapidParam, lngRow:=i + j, _
                               dblBaseX:=rpdDPoint3.x, dblBaseY:=rpdDPoint3.y, _
                               dblShift:=tt, dblAngle:=90
              Set rpdDimeData = Nothing
              Set objRapidTol = Nothing
            End If
            
            If Left(ActiveSheet.Name, 1) = "G" And _
               .blnAsp(j) Then
              objRapidParam.SetUp strParam:="DimeParam"
              rpdDoc.DimensionParam.Height = objRapidParam.dblHW(i, 1) / 1.5
              rpdDoc.DimensionParam.Width = objRapidParam.dblHW(i, 2) / 1.5
              
              Set rpdDimeData = New zwDrawCAD.DimensionData
              rpdDimeData.CreateSingle doc:=rpdDoc, _
                                       PickStart:=rpdDPoint1, PickEnd:=rpdDPoint2, _
                                       PassPoint:=rpdDPoint3, dAngle:=90
              
              rpdDimeData.Text(0) = "(非球面形状範囲)"
              rpdDPoint4.x = rpdDPoint3.x + hh
              rpdDPoint4.y = rpdDPoint3.y
              rpdDimeData.SetTextPointEx doc:=rpdDoc, lIndex:=0, _
                                         PutPoint:=rpdDPoint4, Arrange:=zwDimTextFree, _
                                         MarkDir:=zwMarkDirectInside
              
              objRapidPrim.SetUp prmObject:=rpdDimeData, lngLayerNo:=CLng(Alpha(13)), _
                                 lngPenStyle:=zwStyleSolid, lngPenWidth:=zwWidthThin
              rpdDoc.DataBase.Add Prim:=rpdDimeData
              Set rpdDimeData = Nothing
              
              
              rpdDPoint2.x = .dblCenterX(j) + .dblCornerZ(2, j)
              rpdDPoint2.y = .dblCenterY(j) + .dblCornerH(2, j)
              rpdDPoint3.y = .dblCenterY(j) + (.dblCornerH(2, j) - .dblSurfaceH(j)) / 2
              
              objRapidParam.SetUp strParam:="DimeParam"
              rpdDoc.DimensionParam.Height = objRapidParam.dblHW(i, 1)
              rpdDoc.DimensionParam.Width = objRapidParam.dblHW(i, 2)
              rpdDoc.DimensionParam.DecimalDigit = objRapidParam.lngVal(i, 1)
              
              Set rpdDimeData = New zwDrawCAD.DimensionData
              rpdDimeData.CreateSingle doc:=rpdDoc, _
                                       PickStart:=rpdDPoint1, PickEnd:=rpdDPoint2, _
                                       PassPoint:=rpdDPoint3, dAngle:=90
              
              rpdDimeData.Text(0) = "(" & rpdDimeData.Text(0) & ")"
              yy = Len(rpdDimeData.Text(0)) * (ww / 4) + aa + ww / 8
              rpdDPoint4.x = rpdDPoint3.x
              rpdDPoint4.y = .dblCenterY(j) + .dblCornerH(2, j) + yy
              rpdDimeData.SetTextPoint doc:=rpdDoc, lIndex:=0, _
                                       PutPoint:=rpdDPoint4, Arrange:=zwDimTextParallel
              
              objRapidPrim.SetUp prmObject:=rpdDimeData, lngLayerNo:=CLng(Alpha(13)), _
                                 lngPenStyle:=zwStyleSolid, lngPenWidth:=zwWidthThin
              rpdDoc.DataBase.Add Prim:=rpdDimeData
              Set rpdDimeData = Nothing
              
              
              objRapidParam.SetUp strParam:="DimeParam"
              rpdDoc.DimensionParam.Height = objRapidParam.dblHW(i, 1) / 1.5
              rpdDoc.DimensionParam.Width = objRapidParam.dblHW(i, 2) / 1.5
              
              Set rpdDimeData = New zwDrawCAD.DimensionData
              rpdDimeData.CreateSingle doc:=rpdDoc, _
                                       PickStart:=rpdDPoint1, PickEnd:=rpdDPoint2, _
                                       PassPoint:=rpdDPoint3, dAngle:=90
              
              rpdDimeData.Text(0) = "(平面部)"
              rpdDPoint4.x = rpdDPoint3.x + hh
              rpdDimeData.SetTextPoint doc:=rpdDoc, lIndex:=0, _
                                       PutPoint:=rpdDPoint4, Arrange:=zwDimTextFree
              
              objRapidPrim.SetUp prmObject:=rpdDimeData, lngLayerNo:=CLng(Alpha(13)), _
                                 lngPenStyle:=zwStyleSolid, lngPenWidth:=zwWidthThin
              rpdDoc.DataBase.Add Prim:=rpdDimeData
              Set rpdDimeData = Nothing
            End If
          End If
        Next
    End Select
  Next
  objRapidParam.SetUp strParam:="DimeParam"
End With
Set rpdDPoint1 = Nothing
Set rpdDPoint2 = Nothing
Set rpdDPoint3 = Nothing
Set rpdDPoint4 = Nothing
Set objRapidPrim = Nothing
Set objRapidParam = Nothing

End Sub


Public Sub AddP(ByRef ppdObject As clsPressPlotData)

Dim rpdDimeData As zwDrawCAD.DimensionData
Dim rpdDPoint1 As zwDrawCAD.DPoint
Dim rpdDPoint2 As zwDrawCAD.DPoint
Dim rpdDPoint3 As zwDrawCAD.DPoint
Dim rpdDPoint4 As zwDrawCAD.DPoint
Dim objRapidPrim As clsRapidPrim
Dim objRapidParam As clsRapidParam
Dim objRapidTol As clsRapidTol
Dim objPressParam As clsPressParam

Dim i As Long, j As Long, bc As Long
Dim aa As Double, xx As Double, yy As Double
Dim hh As Double, ww As Double, tt As Double
Dim d(2) As Double

Set objPressParam = New clsPressParam
objPressParam.Import strConfig:="PressParam"
With objPressParam
  For i = 1 To .lngMaxCount("HP寸法")
    If Cells(8, 6).Value < .dbldHP(i, 9) Then
      d(1) = .dbldHP(i, ppdObject.lngZ)
      If d(1) < 0 Then d(1) = 0
      d(2) = .dbldHP(i, ppdObject.lngD)
      If d(2) < 0 Then d(2) = 0
      Exit For
    End If
  Next
End With
Set objPressParam = Nothing

Set rpdDPoint1 = New zwDrawCAD.DPoint
Set rpdDPoint2 = New zwDrawCAD.DPoint
Set rpdDPoint3 = New zwDrawCAD.DPoint
Set rpdDPoint4 = New zwDrawCAD.DPoint
Set objRapidPrim = New clsRapidPrim
Set objRapidParam = New clsRapidParam
objRapidParam.Import strConfig:="Param"
aa = objRapidParam.dblXY(1, 1) / rpdDoc.DrawInfo.Numerator
With ppdObject
  For i = 1 To objRapidParam.lngMaxCount
    Select Case objRapidParam.strItem(i)
      Case "肉厚寸法"
        objProgCtrl.AddBar lngDivCount:=3
        yy = objRapidParam.dblXY(i, 2) / rpdDoc.DrawInfo.Numerator
        hh = objRapidParam.dblHW(i, 1) / rpdDoc.DrawInfo.Numerator
        ww = objRapidParam.dblHW(i, 2) / rpdDoc.DrawInfo.Numerator
        bc = objRapidParam.lngVal(i, 2)
        
        rpdDPoint1.x = .dblCenterX(1)
        rpdDPoint1.y = .dblCenterY(1)
        rpdDPoint2.x = .dblCenterX(2)
        rpdDPoint2.y = .dblCenterY(2)
        rpdDPoint3.x = (.dblCenterX(1) + .dblCenterX(2)) / 2
        rpdDPoint3.y = .dblCenterY(1) - .dblCornerH(2, 1) - yy
        
        objRapidParam.SetUp strParam:="DimeParam"
        rpdDoc.DimensionParam.Height = objRapidParam.dblHW(i, 1)
        rpdDoc.DimensionParam.Width = objRapidParam.dblHW(i, 2)
        rpdDoc.DimensionParam.DecimalDigit = 2
        
        Set rpdDimeData = New zwDrawCAD.DimensionData
        Set objRapidTol = New clsRapidTol
        rpdDimeData.CreateSingle doc:=rpdDoc, _
                                 PickStart:=rpdDPoint1, PickEnd:=rpdDPoint2, _
                                 PassPoint:=rpdDPoint3, dAngle:=0
        
        rpdDimeData.Text(0) = "D.P." & rpdDimeData.Text(0)
        xx = (Len(rpdDimeData.Text(0)) + bc) * (ww / 4) + aa + ww / 8
        tt = (Len(rpdDimeData.Text(0)) - bc) * (ww / 4)
        rpdDimeData.Text(0) = rpdDimeData.Text(0) & Space(bc)
        If (.dblCenterX(2) - .dblCenterX(1)) / 2 < xx Then
          rpdDPoint3.x = .dblCenterX(2) + xx
        End If
        rpdDPoint3.y = rpdDPoint3.y + 0.001 / rpdDoc.DrawInfo.Numerator
        rpdDimeData.SetTextPoint doc:=rpdDoc, lIndex:=0, _
                                 PutPoint:=rpdDPoint3, Arrange:=zwDimTextParallel
        
        objRapidPrim.SetUp prmObject:=rpdDimeData, lngLayerNo:=lngPre(4), _
                           lngPenStyle:=zwStyleSolid, lngPenWidth:=zwWidthThin
        rpdDoc.DataBase.Add Prim:=rpdDimeData
        objRapidTol.AddP rpObject:=objRapidParam, lngRow:=i + 1, _
                         dblBaseX:=rpdDPoint3.x, dblBaseY:=rpdDPoint3.y, _
                         dblShift:=tt, dblAngle:=0, strType:="DP", lngZD:=.lngD
        Set rpdDimeData = Nothing
        Set objRapidTol = Nothing
        
        
        Set rpdDimeData = New zwDrawCAD.DimensionData
        Set objRapidTol = New clsRapidTol
        rpdDimeData.CreateSingle doc:=rpdDoc, _
                                 PickStart:=rpdDPoint1, PickEnd:=rpdDPoint2, _
                                 PassPoint:=rpdDPoint3, dAngle:=0
        
        rpdDimeData.Text(0) = Format(CDbl(rpdDimeData.Text(0)) + d(2), "0.00")
        rpdDimeData.Text(0) = "H.P." & rpdDimeData.Text(0) & Space(bc)
        rpdDPoint4.x = rpdDPoint3.x
        rpdDPoint4.y = rpdDPoint3.y + hh * 1.2
        rpdDimeData.SetTextPoint doc:=rpdDoc, lIndex:=0, _
                                 PutPoint:=rpdDPoint4, Arrange:=zwDimTextFree
        
        objRapidPrim.SetUp prmObject:=rpdDimeData, lngLayerNo:=lngPre(4), _
                           lngPenStyle:=zwStyleSolid, lngPenWidth:=zwWidthThin
        rpdDoc.DataBase.Add Prim:=rpdDimeData
        objRapidTol.AddP rpObject:=objRapidParam, lngRow:=i + 1, _
                         dblBaseX:=rpdDPoint4.x, dblBaseY:=rpdDPoint4.y, _
                         dblShift:=tt, dblAngle:=0, strType:="HP", lngZD:=.lngD
        Set rpdDimeData = Nothing
        Set objRapidTol = Nothing
        
      Case "外径寸法"
        objProgCtrl.AddBar lngDivCount:=3
        xx = objRapidParam.dblXY(i, 1) / rpdDoc.DrawInfo.Numerator
        hh = objRapidParam.dblHW(i, 1) / rpdDoc.DrawInfo.Numerator
        ww = objRapidParam.dblHW(i, 2) / rpdDoc.DrawInfo.Numerator
        bc = objRapidParam.lngVal(i, 2)
        
        rpdDPoint1.x = .dblCenterX(1) + .dblCornerZ(2, 1)
        rpdDPoint1.y = .dblCenterY(1) + .dblCornerH(2, 1)
        rpdDPoint2.x = .dblCenterX(1) + .dblCornerZ(2, 1)
        rpdDPoint2.y = .dblCenterY(1) - .dblCornerH(2, 1)
        If .dblCornerZ(1, 1) < 0 Then
          rpdDPoint3.x = .dblCenterX(1) + .dblCornerZ(1, 1) - xx
        Else
          rpdDPoint3.x = .dblCenterX(1) - xx
        End If
        rpdDPoint3.y = .dblCenterY(1)
        
        objRapidParam.SetUp strParam:="DimeParam"
        rpdDoc.DimensionParam.Height = objRapidParam.dblHW(i, 1)
        rpdDoc.DimensionParam.Width = objRapidParam.dblHW(i, 2)
        rpdDoc.DimensionParam.DecimalDigit = 1
        
        Set rpdDimeData = New zwDrawCAD.DimensionData
        Set objRapidTol = New clsRapidTol
        rpdDimeData.CreateSingle doc:=rpdDoc, _
                                 PickStart:=rpdDPoint1, PickEnd:=rpdDPoint2, _
                                 PassPoint:=rpdDPoint3, dAngle:=90
        
        rpdDimeData.Text(0) = "D.P.φ" & rpdDimeData.Text(0)
        tt = (Len(rpdDimeData.Text(0)) + 1 - bc) * (ww / 4)
        rpdDimeData.Text(0) = rpdDimeData.Text(0) & Space(bc)
        
        objRapidPrim.SetUp prmObject:=rpdDimeData, lngLayerNo:=lngPre(4), _
                           lngPenStyle:=zwStyleSolid, lngPenWidth:=zwWidthThin
        rpdDoc.DataBase.Add Prim:=rpdDimeData
        objRapidTol.AddP rpObject:=objRapidParam, lngRow:=i + 1, _
                         dblBaseX:=rpdDPoint3.x, dblBaseY:=rpdDPoint3.y, _
                         dblShift:=tt, dblAngle:=90, strType:="DP", lngZD:=.lngZ
        Set rpdDimeData = Nothing
        Set objRapidTol = Nothing
        
        
        Set rpdDimeData = New zwDrawCAD.DimensionData
        Set objRapidTol = New clsRapidTol
        rpdDimeData.CreateSingle doc:=rpdDoc, _
                                 PickStart:=rpdDPoint1, PickEnd:=rpdDPoint2, _
                                 PassPoint:=rpdDPoint3, dAngle:=90
        
        rpdDimeData.Text(0) = Format(CDbl(rpdDimeData.Text(0)) + d(1), "0.0")
        rpdDimeData.Text(0) = "H.P.φ" & rpdDimeData.Text(0) & Space(bc)
        rpdDPoint4.x = rpdDPoint3.x - hh * 1.2
        rpdDPoint4.y = rpdDPoint3.y
        rpdDimeData.SetTextPointEx doc:=rpdDoc, lIndex:=0, _
                                   PutPoint:=rpdDPoint4, Arrange:=zwDimTextFree, _
                                   MarkDir:=zwMarkDirectInside
        
        objRapidPrim.SetUp prmObject:=rpdDimeData, lngLayerNo:=lngPre(4), _
                           lngPenStyle:=zwStyleSolid, lngPenWidth:=zwWidthThin
        rpdDoc.DataBase.Add Prim:=rpdDimeData
        objRapidTol.AddP rpObject:=objRapidParam, lngRow:=i + 1, _
                         dblBaseX:=rpdDPoint4.x, dblBaseY:=rpdDPoint4.y, _
                         dblShift:=tt, dblAngle:=90, strType:="HP", lngZD:=.lngZ
        Set rpdDimeData = Nothing
        Set objRapidTol = Nothing
        
      Case "内径寸法"
        objProgCtrl.AddBar lngDivCount:=3
        xx = objRapidParam.dblXY(i, 1) / rpdDoc.DrawInfo.Numerator
        hh = objRapidParam.dblHW(i, 1) / rpdDoc.DrawInfo.Numerator
        
        For j = 1 To 2
          If .dblCornerH(1, j) - .dblSurfaceH(j) > 0 Then
            rpdDPoint1.x = .dblCenterX(j) + .dblCornerZ(1, j)
            rpdDPoint1.y = .dblCenterY(j) + .dblSurfaceH(j)
            rpdDPoint2.x = .dblCenterX(j) + .dblCornerZ(1, j)
            rpdDPoint2.y = .dblCenterY(j) - .dblSurfaceH(j)
            If .dblCornerZ(1, j) * (j * (-2) + 3) < 0 Then
              rpdDPoint3.x = .dblCenterX(j) + .dblCornerZ(1, j) + (xx - hh) * (j * 2 - 3)
            Else
              rpdDPoint3.x = .dblCenterX(j) + (xx - hh) * (j * 2 - 3)
            End If
            rpdDPoint3.y = .dblCenterY(j)
            
            objRapidParam.SetUp strParam:="DimeParam"
            rpdDoc.DimensionParam.Height = objRapidParam.dblHW(i, 1)
            rpdDoc.DimensionParam.Width = objRapidParam.dblHW(i, 2)
            rpdDoc.DimensionParam.DecimalDigit = 1
            
            Set rpdDimeData = New zwDrawCAD.DimensionData
            rpdDimeData.CreateSingle doc:=rpdDoc, _
                                     PickStart:=rpdDPoint1, PickEnd:=rpdDPoint2, _
                                     PassPoint:=rpdDPoint3, dAngle:=90
            
            rpdDimeData.Text(0) = "D.P.φ" & rpdDimeData.Text(0)
            
            objRapidPrim.SetUp prmObject:=rpdDimeData, lngLayerNo:=lngPre(4), _
                               lngPenStyle:=zwStyleSolid, lngPenWidth:=zwWidthThin
            rpdDoc.DataBase.Add Prim:=rpdDimeData
            Set rpdDimeData = Nothing
            
            
            Set rpdDimeData = New zwDrawCAD.DimensionData
            rpdDimeData.CreateSingle doc:=rpdDoc, _
                                     PickStart:=rpdDPoint1, PickEnd:=rpdDPoint2, _
                                     PassPoint:=rpdDPoint3, dAngle:=90
            
            rpdDimeData.Text(0) = "H.P.φ" & rpdDimeData.Text(0)
            rpdDPoint4.x = rpdDPoint3.x - hh
            rpdDPoint4.y = rpdDPoint3.y
            rpdDimeData.SetTextPointEx doc:=rpdDoc, lIndex:=0, _
                                       PutPoint:=rpdDPoint4, Arrange:=zwDimTextFree, _
                                       MarkDir:=zwMarkDirectInside
            
            objRapidPrim.SetUp prmObject:=rpdDimeData, lngLayerNo:=lngPre(4), _
                               lngPenStyle:=zwStyleSolid, lngPenWidth:=zwWidthThin
            rpdDoc.DataBase.Add Prim:=rpdDimeData
            Set rpdDimeData = Nothing
          End If
        Next
    End Select
  Next
  objRapidParam.SetUp strParam:="DimeParam"
End With
Set rpdDPoint1 = Nothing
Set rpdDPoint2 = Nothing
Set rpdDPoint3 = Nothing
Set rpdDPoint4 = Nothing
Set objRapidPrim = Nothing
Set objRapidParam = Nothing

End Sub


Public Sub AddB(ByRef bpdObject As clsBalsamPlotData)

Dim rpdDimeData As zwDrawCAD.DimensionData
Dim rpdDPoint1 As zwDrawCAD.DPoint
Dim rpdDPoint2 As zwDrawCAD.DPoint
Dim rpdDPoint3 As zwDrawCAD.DPoint
Dim rpdDPoint4 As zwDrawCAD.DPoint
Dim objRapidPrim As clsRapidPrim
Dim objRapidParam As clsRapidParam
Dim objRapidTol As clsRapidTol

Dim i As Long, j As Long, m As Long
Dim dd As Long, bc As Long
Dim aa As Double, xx As Double, yy As Double
Dim hh As Double, ww As Double, tt As Double
Dim asName As String, prtName(2) As String

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

Set rpdDPoint1 = New zwDrawCAD.DPoint
Set rpdDPoint2 = New zwDrawCAD.DPoint
Set rpdDPoint3 = New zwDrawCAD.DPoint
Set rpdDPoint4 = New zwDrawCAD.DPoint
Set objRapidPrim = New clsRapidPrim
Set objRapidParam = New clsRapidParam
objRapidParam.Import strConfig:="Param"
aa = objRapidParam.dblXY(1, 1) / rpdDoc.DrawInfo.Numerator
With bpdObject
  For i = 1 To objRapidParam.lngMaxCount
    Select Case objRapidParam.strItem(i)
      Case "肉厚寸法"
        objProgCtrl.AddBar lngDivCount:=3
        yy = objRapidParam.dblXY(i, 2) / rpdDoc.DrawInfo.Numerator
        ww = objRapidParam.dblHW(i, 2) / rpdDoc.DrawInfo.Numerator
        bc = objRapidParam.lngVal(i, 2)
        
        rpdDPoint1.x = .dblCenterX(1)
        rpdDPoint1.y = .dblCenterY(1)
        rpdDPoint2.x = .dblCenterX(2)
        rpdDPoint2.y = .dblCenterY(2)
        rpdDPoint3.x = (.dblCenterX(1) + .dblCenterX(2)) / 2
        rpdDPoint3.y = .dblCenterY(1) - .dblMaxDiameter / 2 - yy
        
        objRapidParam.SetUp strParam:="DimeParam"
        rpdDoc.DimensionParam.Height = objRapidParam.dblHW(i, 1)
        rpdDoc.DimensionParam.Width = objRapidParam.dblHW(i, 2)
        rpdDoc.DimensionParam.DecimalDigit = objRapidParam.lngVal(i, 1)
        
        Set rpdDimeData = New zwDrawCAD.DimensionData
        Set objRapidTol = New clsRapidTol
        rpdDimeData.CreateSingle doc:=rpdDoc, _
                                 PickStart:=rpdDPoint1, PickEnd:=rpdDPoint2, _
                                 PassPoint:=rpdDPoint3, dAngle:=0
        
        xx = (Len(rpdDimeData.Text(0)) + bc) * (ww / 4) + aa + ww / 8
        tt = (Len(rpdDimeData.Text(0)) - bc) * (ww / 4)
        rpdDimeData.Text(0) = rpdDimeData.Text(0) & Space(bc)
        If (.dblCenterX(2) - .dblCenterX(1)) / 2 < xx Then
          rpdDPoint3.x = .dblCenterX(2) + xx
        End If
        rpdDPoint3.y = rpdDPoint3.y + 0.001 / rpdDoc.DrawInfo.Numerator
        rpdDimeData.SetTextPoint doc:=rpdDoc, lIndex:=0, _
                                 PutPoint:=rpdDPoint3, Arrange:=zwDimTextParallel
        
        objRapidPrim.SetUp prmObject:=rpdDimeData, lngLayerNo:=CLng(Alpha(13)), _
                           lngPenStyle:=zwStyleSolid, lngPenWidth:=zwWidthThin
        rpdDoc.DataBase.Add Prim:=rpdDimeData
        objRapidTol.AddB rpObject:=objRapidParam, lngRow:=i + 1, _
                         dblBaseX:=rpdDPoint3.x, dblBaseY:=rpdDPoint3.y, _
                         dblShift:=tt, dblAngle:=0
        Set rpdDimeData = Nothing
        Set objRapidTol = Nothing
        
      Case "⊿Ｈ寸法"
        objProgCtrl.AddBar lngDivCount:=3
        yy = objRapidParam.dblXY(i, 2) / rpdDoc.DrawInfo.Numerator
        ww = objRapidParam.dblHW(i, 2) / rpdDoc.DrawInfo.Numerator
        dd = objRapidParam.lngVal(i, 1)
        bc = objRapidParam.lngVal(i, 2)
        
        For j = 1 To 2
          If .dblBalsamH <> 0 And _
             Left(ActiveSheet.Name, 1) = "M" And _
             Left(prtName(j), 1) = "H" Then
            rpdDPoint1.x = .dblCenterX(j)
            rpdDPoint1.y = .dblCenterY(j)
            rpdDPoint2.x = .dblCenterX(j) + .dblBalsamZ
            rpdDPoint2.y = .dblCenterY(j) - .dblBalsamH
            rpdDPoint3.x = .dblCenterX(j) + .dblBalsamZ / 2
            rpdDPoint3.y = .dblCenterY(1) - .dblMaxDiameter / 2 - yy
            
            objRapidParam.SetUp strParam:="DimeParam"
            rpdDoc.DimensionParam.Height = objRapidParam.dblHW(i, 1)
            rpdDoc.DimensionParam.Width = objRapidParam.dblHW(i, 2)
            rpdDoc.DimensionParam.DecimalDigit = 0
            If .dblBalsamZ * (j * (-2) + 3) < 0 Then
              rpdDoc.DimensionParam.StartMarkType = zwMarkBlackCircle
              rpdDoc.DimensionParam.StartDistance = 0.001 / rpdDoc.DrawInfo.Numerator
            Else
              rpdDoc.DimensionParam.EndMarkType = zwMarkBlackCircle
              rpdDoc.DimensionParam.EndDistance = 0.001 / rpdDoc.DrawInfo.Numerator
            End If
            
            Set rpdDimeData = New zwDrawCAD.DimensionData
            Set objRapidTol = New clsRapidTol
            rpdDimeData.CreateSingle doc:=rpdDoc, _
                                     PickStart:=rpdDPoint1, PickEnd:=rpdDPoint2, _
                                     PassPoint:=rpdDPoint3, dAngle:=0
            
            rpdDimeData.Text(0) = Format(Abs(.dblBalsamZ), _
                                         "0." & String(dd, "0"))
            xx = (Len(rpdDimeData.Text(0)) + bc) * (ww / 4) + aa + ww / 8
            tt = (Len(rpdDimeData.Text(0)) - bc) * (ww / 4)
            rpdDimeData.Text(0) = rpdDimeData.Text(0) & Space(bc)
            If .dblBalsamZ * (j * (-2) + 3) < 0 Then
              rpdDPoint3.x = .dblCenterX(j) + .dblBalsamZ + xx * (j * 2 - 3)
            Else
              rpdDPoint3.x = .dblCenterX(j) + xx * (j * 2 - 3)
            End If
            rpdDPoint3.y = rpdDPoint3.y + 0.001 / rpdDoc.DrawInfo.Numerator
            rpdDimeData.SetTextPoint doc:=rpdDoc, lIndex:=0, _
                                     PutPoint:=rpdDPoint3, Arrange:=zwDimTextParallel
            
            objRapidPrim.SetUp prmObject:=rpdDimeData, lngLayerNo:=CLng(Alpha(13)), _
                               lngPenStyle:=zwStyleSolid, lngPenWidth:=zwWidthThin
            rpdDoc.DataBase.Add Prim:=rpdDimeData
            objRapidTol.AddB rpObject:=objRapidParam, lngRow:=i + j, _
                             dblBaseX:=rpdDPoint3.x, dblBaseY:=rpdDPoint3.y, _
                             dblShift:=tt, dblAngle:=0
            Set rpdDimeData = Nothing
            Set objRapidTol = Nothing
          End If
        Next
        
      Case "外径寸法"
        objProgCtrl.AddBar lngDivCount:=3
        xx = objRapidParam.dblXY(i, 1) / rpdDoc.DrawInfo.Numerator
        hh = objRapidParam.dblHW(i, 1) / rpdDoc.DrawInfo.Numerator
        ww = objRapidParam.dblHW(i, 2) / rpdDoc.DrawInfo.Numerator
        bc = objRapidParam.lngVal(i, 2)
        
        If .dblCornerH(1) > .dblCornerH(2) Then
          m = 2
        Else
          m = 1
        End If
        
        For j = 1 To 2
          If Left(ActiveSheet.Name, 1) = "B" Or _
             Left(ActiveSheet.Name, 1) = "M" And _
             j <> m Then
            rpdDPoint1.x = .dblCenterX(j) + .dblCornerZ(j)
            rpdDPoint1.y = .dblCenterY(j) + .dblCornerH(j)
            rpdDPoint2.x = .dblCenterX(j) + .dblCornerZ(j)
            rpdDPoint2.y = .dblCenterY(j) - .dblCornerH(j)
            If j = m Then
              rpdDPoint3.x = .dblCenterX(j) + (xx - hh * 3) * (j * 2 - 3)
            Else
              rpdDPoint3.x = .dblCenterX(j) + xx * (j * 2 - 3)
            End If
            If .dblCornerZ(j) * (j * (-2) + 3) < 0 Then
              rpdDPoint3.x = rpdDPoint3.x + .dblCornerZ(j)
            End If
            rpdDPoint3.y = .dblCenterY(j)
            
            objRapidParam.SetUp strParam:="DimeParam"
            rpdDoc.DimensionParam.Height = objRapidParam.dblHW(i, 1)
            rpdDoc.DimensionParam.Width = objRapidParam.dblHW(i, 2)
            rpdDoc.DimensionParam.DecimalDigit = objRapidParam.lngVal(i, 1)
            
            Worksheets(prtName(j)).Activate
            Set rpdDimeData = New zwDrawCAD.DimensionData
            Set objRapidTol = New clsRapidTol
            rpdDimeData.CreateSingle doc:=rpdDoc, _
                                     PickStart:=rpdDPoint1, PickEnd:=rpdDPoint2, _
                                     PassPoint:=rpdDPoint3, dAngle:=90
            
            rpdDimeData.Text(0) = "φ" & rpdDimeData.Text(0)
            tt = 0
            objRapidTol.Check rpObject:=objRapidParam, lngRow:=i + 1
            If objRapidTol.strKeyWord = "Num" Then
              tt = (Len(rpdDimeData.Text(0)) + 1 - bc) * (ww / 4)
              rpdDimeData.Text(0) = rpdDimeData.Text(0) & Space(bc)
            ElseIf objRapidTol.strKeyWord <> "Empty" Then
              rpdDimeData.Text(0) = rpdDimeData.Text(0) & objRapidTol.strKeyWord
            End If
            
            objRapidPrim.SetUp prmObject:=rpdDimeData, lngLayerNo:=CLng(Alpha(13)), _
                               lngPenStyle:=zwStyleSolid, lngPenWidth:=zwWidthThin
            rpdDoc.DataBase.Add Prim:=rpdDimeData
            objRapidTol.AddG rpObject:=objRapidParam, lngRow:=i + 1, _
                             dblBaseX:=rpdDPoint3.x, dblBaseY:=rpdDPoint3.y, _
                             dblShift:=tt, dblAngle:=90
            Set rpdDimeData = Nothing
            Set objRapidTol = Nothing
            Worksheets(asName).Activate
          End If
        Next
        
        If .dblBalsamH <> 0 And _
           Left(ActiveSheet.Name, 1) = "B" Then
          rpdDPoint1.x = .dblCenterX(m) + .dblBalsamZ
          rpdDPoint1.y = .dblCenterY(m) + .dblBalsamH
          rpdDPoint2.x = .dblCenterX(m) + .dblBalsamZ
          rpdDPoint2.y = .dblCenterY(m) - .dblBalsamH
          If .dblCornerZ(m) * (m * (-2) + 3) < 0 Then
            rpdDPoint3.x = .dblCenterX(m) + .dblCornerZ(m) + xx * (m * 2 - 3)
          Else
            rpdDPoint3.x = .dblCenterX(m) + xx * (m * 2 - 3)
          End If
          rpdDPoint3.y = .dblCenterY(m)
          
          objRapidParam.SetUp strParam:="DimeParam"
          rpdDoc.DimensionParam.Height = objRapidParam.dblHW(i, 1)
          rpdDoc.DimensionParam.Width = objRapidParam.dblHW(i, 2)
          rpdDoc.DimensionParam.DecimalDigit = 1
          
          Set rpdDimeData = New zwDrawCAD.DimensionData
          rpdDimeData.CreateSingle doc:=rpdDoc, _
                                   PickStart:=rpdDPoint1, PickEnd:=rpdDPoint2, _
                                   PassPoint:=rpdDPoint3, dAngle:=90
          
          rpdDimeData.Text(0) = "φ" & rpdDimeData.Text(0) & "以下"
          
          objRapidPrim.SetUp prmObject:=rpdDimeData, lngLayerNo:=CLng(Alpha(13)), _
                             lngPenStyle:=zwStyleSolid, lngPenWidth:=zwWidthThin
          rpdDoc.DataBase.Add Prim:=rpdDimeData
          Set rpdDimeData = Nothing
          
          objRapidParam.SetUp strParam:="DimeParam"
          rpdDoc.DimensionParam.Height = objRapidParam.dblHW(i, 1) / 1.5
          rpdDoc.DimensionParam.Width = objRapidParam.dblHW(i, 2) / 2
          
          Set rpdDimeData = New zwDrawCAD.DimensionData
          rpdDimeData.CreateSingle doc:=rpdDoc, _
                                   PickStart:=rpdDPoint1, PickEnd:=rpdDPoint2, _
                                   PassPoint:=rpdDPoint3, dAngle:=90
          
          rpdDimeData.Text(0) = "(接着剤はみ出し許容範囲)"
          rpdDPoint4.x = rpdDPoint3.x + hh
          rpdDPoint4.y = rpdDPoint3.y
          rpdDimeData.SetTextPointEx doc:=rpdDoc, lIndex:=0, _
                                     PutPoint:=rpdDPoint4, Arrange:=zwDimTextFree, _
                                     MarkDir:=zwMarkDirectInside
          
          objRapidPrim.SetUp prmObject:=rpdDimeData, lngLayerNo:=CLng(Alpha(13)), _
                             lngPenStyle:=zwStyleSolid, lngPenWidth:=zwWidthThin
          rpdDoc.DataBase.Add Prim:=rpdDimeData
          Set rpdDimeData = Nothing
        End If
    End Select
  Next
  objRapidParam.SetUp strParam:="DimeParam"
End With
Set rpdDPoint1 = Nothing
Set rpdDPoint2 = Nothing
Set rpdDPoint3 = Nothing
Set rpdDPoint4 = Nothing
Set objRapidPrim = Nothing
Set objRapidParam = Nothing

End Sub


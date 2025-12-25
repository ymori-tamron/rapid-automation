VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsAssyPlotData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False


Option Explicit
Option Base 1
Option Compare Binary

Private varOutE(99, 10) As Variant
Private lngOutMC As Long
Private dblOutMD As Double, dblOutML As Double
Private dblOutCX As Double, dblOutCY As Double


Private Sub Class_Initialize()

Dim i As Long, j As Long

For i = 1 To UBound(varOutE, 1)
  For j = 1 To UBound(varOutE, 2)
    Select Case j
      Case 1, 8, 9, 10
        varOutE(i, j) = ""
      Case Else
        varOutE(i, j) = 0#
    End Select
  Next
Next
lngOutMC = 0
dblOutMD = 0: dblOutML = 0
dblOutCX = 0: dblOutCY = 0

End Sub


Public Property Get strEName(ByVal lngIndex As Long) As String

strEName = varOutE(lngIndex, 1)

End Property


Public Property Get lngMaxCount() As Long

lngMaxCount = lngOutMC

End Property


Public Property Get dblMaxDiameter() As Double

dblMaxDiameter = dblOutMD

End Property


Public Property Get dblMaxLength() As Double

dblMaxLength = dblOutML

End Property


Public Property Get dblCenterX() As Double

dblCenterX = dblOutCX

End Property


Public Property Get dblCenterY() As Double

dblCenterY = dblOutCY

End Property


Public Sub Calc()

Dim xlWs As Excel.Worksheet

Dim k As Variant
Dim chk As Boolean
Dim i As Long, j As Long, pos As Long
Dim maxRow(2) As Long
Dim txt As String

If Not IsEmpty(Cells(8, 6).Value) Then
  pos = CLng(Mid(Cells(8, 6).Value, 4))
Else
  pos = 0
End If

With Worksheets(DefaultSheet1)
  chk = False
  maxRow(1) = .Cells(6, 4).End(xlDown).Row
  If maxRow(1) > 6 Then
    For i = 7 To maxRow(1)
      If .Cells(i, 3).Value = Left(Cells(8, 2).Value, 3) Then chk = True
      If chk Then
        If Not IsEmpty(.Cells(i, 3).Value) Then
          lngOutMC = lngOutMC + 1
          varOutE(lngOutMC, 1) = .Cells(i, 3).Value
          varOutE(lngOutMC, 2) = dblOutML
          varOutE(lngOutMC, 8) = .Cells(i, 10).Value
          varOutE(lngOutMC, 10) = .Cells(i, 2).Value
          
          If .Cells(i, 3).Value = "STO" Or _
             Left(.Cells(i, 3).Value, 2) = "AP" Then
            varOutE(lngOutMC, 3) = .Cells(i, 9).Value * 2
          Else
            If .Cells(i, 3).Value = "HYB" Then
              If Left(.Cells(i - 1, 3).Value, 1) = "G" Then
                varOutE(lngOutMC, 1) = "H" & Mid(.Cells(i - 1, 3).Value, 2)
              ElseIf Left(.Cells(i + 1, 3).Value, 1) = "G" Then
                varOutE(lngOutMC, 1) = "H" & Mid(.Cells(i + 1, 3).Value, 2)
              End If
            End If
            If InStr(.Cells(i, 4).Value, "A") > 0 Or _
               InStr(.Cells(i + 1, 4).Value, "A") > 0 Then
              varOutE(lngOutMC, 1) = varOutE(lngOutMC, 1) & "(ASP)"
            End If
            varOutE(lngOutMC, 9) = .Cells(i + 1, 10).Value
            
            Set xlWs = Worksheets(varOutE(lngOutMC, 1))
            varOutE(lngOutMC, 3) = xlWs.Cells(8, 6).Value
            If xlWs.Cells(8, 6).Value > dblOutMD Then
              dblOutMD = xlWs.Cells(8, 6).Value
            End If
            
            varOutE(lngOutMC, 4) = xlWs.Cells(13, 3).Value
            varOutE(lngOutMC, 5) = xlWs.Cells(14, 3).Value
            
            For Each k In Worksheets
              If k.Name = "M" & Mid(xlWs.Name, 2, 2) Then
                varOutE(lngOutMC, 4) = k.Cells(10, 3).Value
                varOutE(lngOutMC, 5) = k.Cells(11, 3).Value
                If .Cells(i - 1, 3).Value = "HYB" Then
                  varOutE(lngOutMC, 8) = .Cells(i - 1, 10).Value
                ElseIf .Cells(i + 1, 3).Value = "HYB" Then
                  varOutE(lngOutMC, 9) = .Cells(i + 2, 10).Value
                End If
                Exit For
              End If
            Next
            Set xlWs = Nothing
          End If
        End If
        
        If IsNumeric(.Cells(i, 6).Value) Then
          dblOutML = dblOutML + .Cells(i, 6).Value
        Else
          If pos > 0 Then
            txt = Mid(.Cells(i, 6).Value, 2, Len(.Cells(i, 6).Value) - 2)
            maxRow(2) = .Cells(maxRow(1) + 3, 4).End(xlDown).Row
            If maxRow(2) > maxRow(1) + 3 Then
              For j = maxRow(1) + 4 To maxRow(2)
                If .Cells(j, 3).Value = txt Then
                  dblOutML = dblOutML + .Cells(j, 4 + pos).Value
                  Exit For
                End If
              Next
            End If
          End If
        End If
      End If
      If .Cells(i, 3).Value = _
        Right(Cells(7, 2).End(xlDown).Value, 3) Then Exit For
    Next
  End If
End With

End Sub


Public Sub CreateElement()

Dim rpdLineData As zwDrawCAD.LineData
Dim rpdPolyData As zwDrawCAD.PolyLineData
Dim rpdDPoints As zwDrawCAD.DPoints
Dim rpdDPoint1 As zwDrawCAD.DPoint
Dim rpdDPoint2 As zwDrawCAD.DPoint
Dim rpdDPoint3 As zwDrawCAD.DPoint
Dim objRapidPrim As clsRapidPrim
Dim objRapidParam As clsRapidParam
Dim objRapidLine As clsRapidLine
Dim objGPData As clsGlassPlotData

Dim k As Variant
Dim i As Long, m As Long
Dim bc As Long, ac As Long
Dim aa As Double
Dim asName As String

Set objRapidParam = New clsRapidParam
objRapidParam.Import strConfig:="Param"
aa = objRapidParam.dblXY(1, 1) / rpdDoc.DrawInfo.Numerator
Set objRapidParam = Nothing

asName = ActiveSheet.Name
dblOutCX = 216 / rpdDoc.DrawInfo.Numerator
dblOutCY = 163 / rpdDoc.DrawInfo.Numerator

Set rpdLineData = New zwDrawCAD.LineData
Set rpdDPoint1 = New zwDrawCAD.DPoint
Set rpdDPoint2 = New zwDrawCAD.DPoint
Set rpdDPoint3 = New zwDrawCAD.DPoint
Set objRapidPrim = New clsRapidPrim
Set objRapidLine = New clsRapidLine
Set objGPData = New clsGlassPlotData
With rpdDoc.DataBase
  For Each k In Worksheets
    If k.Name <> DefaultSheet1 And _
       k.Name <> DefaultSheet3 Then
      If k.Name <> "G00" And _
         Left(k.Name, 1) <> "B" And _
         Left(k.Name, 1) <> "M" Then
        k.Activate: DoEvents
        If varElement(ActiveSheet.Name, "POS") <> -1 Then
          bc = .Count
          objGPData.Calc
          objRapidLine.AddG gpdObject:=objGPData
          If Left(ActiveSheet.Name, 1) <> "H" Then
            .Item(.Count - 1).Selected = True
            .RemoveSelect
          End If
          ac = .Count - 1
          For i = bc To ac
            .Item(i).Selected = True
          Next
          .MoveS _
            dX:=dblOutCX - dblOutML / 2 + _
                varElement(ActiveSheet.Name, "POS") - _
                objGPData.dblCenterX(1), _
            dY:=dblOutCY - objGPData.dblCenterY(1), _
            KeepSelect:=zwKeepSelectNon
          rpdDoc.Views(0).Redraw
          
          For i = 1 To lngOutMC
            If varOutE(i, 1) = ActiveSheet.Name Then
              varOutE(i, 6) = (objGPData.dblCornerZ(2, 1) + _
                               objGPData.dblCornerZ(2, 2) + _
                               Cells(10, 3).Value) / 2
              varOutE(i, 7) = objGPData.dblCornerH(2, 1) * 0.9
            End If
          Next
        Else
          objProgCtrl.AddBar lngDivCount:=1
        End If
      End If
    End If
  Next
  Worksheets(asName).Activate
End With

With rpdDoc
  If varElement("STO", "POS") <> -1 Then
    For m = -1 To 1 Step 2
      objProgCtrl.AddBar lngDivCount:=2
      rpdDPoint1.x = dblOutCX - dblOutML / 2 + varElement("STO", "POS")
      rpdDPoint1.y = dblOutCY + (varElement("STO", "DIA") / 2) * m
      rpdDPoint2.x = dblOutCX - dblOutML / 2 + varElement("STO", "POS")
      rpdDPoint2.y = dblOutCY + (varElement("STO", "DIA") / 2 + aa * 2) * m
      
      Set rpdPolyData = New zwDrawCAD.PolyLineData
      rpdPolyData.CreateArrow Type:=zwMarkBlackArrow, _
                              StartPoint:=rpdDPoint1, EndPoint:=rpdDPoint2, _
                              dSize:=aa
      objRapidPrim.SetUp prmObject:=rpdPolyData, lngLayerNo:=CLng(Alpha(10)), _
                         lngPenStyle:=zwStyleSolid, lngPenWidth:=zwWidthMidThin
      .DataBase.Add Prim:=rpdPolyData
      Set rpdPolyData = Nothing
      
      rpdLineData.Start.x = dblOutCX - dblOutML / 2 + varElement("STO", "POS")
      rpdLineData.Start.y = dblOutCY
      rpdLineData.End.x = dblOutCX - dblOutML / 2 + varElement("STO", "POS")
      rpdLineData.End.y = dblOutCY + (varElement("STO", "DIA") / 2) * m
      objRapidPrim.SetUp prmObject:=rpdLineData, lngLayerNo:=CLng(Alpha(13)), _
                       lngPenStyle:=zwStyleSolid, lngPenWidth:=zwWidthThin
      .DataBase.Add Prim:=rpdLineData
    Next
  End If
  
  For i = 1 To lngOutMC
    If varElement("AP" & CStr(i), "POS") <> -1 Then
      For m = -1 To 1 Step 2
        objProgCtrl.AddBar lngDivCount:=2
        rpdDPoint1.x = dblOutCX - dblOutML / 2 + varElement("AP" & CStr(i), "POS")
        rpdDPoint1.y = dblOutCY + (varElement("AP" & CStr(i), "DIA") / 2) * m
        rpdDPoint2.x = dblOutCX - dblOutML / 2 + varElement("AP" & CStr(i), "POS")
        rpdDPoint2.y = dblOutCY + (varElement("AP" & CStr(i), "DIA") / 2 + aa * 2) * m
        
        Set rpdPolyData = New zwDrawCAD.PolyLineData
        rpdPolyData.CreateArrow Type:=zwMarkBlackArrow, _
                                StartPoint:=rpdDPoint1, EndPoint:=rpdDPoint2, _
                                dSize:=aa
        objRapidPrim.SetUp prmObject:=rpdPolyData, lngLayerNo:=CLng(Alpha(10)), _
                           lngPenStyle:=zwStyleSolid, lngPenWidth:=zwWidthMidThin
        .DataBase.Add Prim:=rpdPolyData
        Set rpdPolyData = Nothing
        
        rpdLineData.Start.x = dblOutCX - dblOutML / 2 + varElement("AP" & CStr(i), "POS")
        rpdLineData.Start.y = dblOutCY
        rpdLineData.End.x = dblOutCX - dblOutML / 2 + varElement("AP" & CStr(i), "POS")
        rpdLineData.End.y = dblOutCY + (varElement("AP" & CStr(i), "DIA") / 2) * m
        objRapidPrim.SetUp prmObject:=rpdLineData, lngLayerNo:=CLng(Alpha(13)), _
                       lngPenStyle:=zwStyleSolid, lngPenWidth:=zwWidthThin
        .DataBase.Add Prim:=rpdLineData
      Next
    Else
      Exit For
    End If
  Next
  
  
  rpdDPoint3.x = dblOutCX + dblOutML / 2 + 20 / .DrawInfo.Numerator
  rpdDPoint3.y = dblOutCY
  
  Set rpdPolyData = New zwDrawCAD.PolyLineData
  rpdPolyData.CreateMark Type:=zwMarkCross, _
                         PutPoint:=rpdDPoint3, dSize:=aa * 2
  objRapidPrim.SetUp prmObject:=rpdPolyData, lngLayerNo:=CLng(Alpha(10)), _
                     lngPenStyle:=zwStyleSolid, lngPenWidth:=zwWidthMidThin
  .DataBase.Add Prim:=rpdPolyData
  Set rpdPolyData = Nothing
  
  For i = 1 To 2
    Set rpdDPoints = New zwDrawCAD.DPoints
    rpdDPoints.Add dX:=dblOutCX + dblOutML / 2 + (8 + i) / .DrawInfo.Numerator, _
                   dY:=dblOutCY + aa * 3
    rpdDPoints.Add dX:=dblOutCX + dblOutML / 2 + (9 + i) / .DrawInfo.Numerator, _
                   dY:=dblOutCY + aa
    rpdDPoints.Add dX:=dblOutCX + dblOutML / 2 + (8 + i) / .DrawInfo.Numerator, _
                   dY:=dblOutCY - aa
    rpdDPoints.Add dX:=dblOutCX + dblOutML / 2 + (9 + i) / .DrawInfo.Numerator, _
                   dY:=dblOutCY - aa * 3
    
    Set rpdPolyData = New zwDrawCAD.PolyLineData
    rpdPolyData.Create Type:=zwPolySpline2, PassPoints:=rpdDPoints
    objRapidPrim.SetUp prmObject:=rpdPolyData, lngLayerNo:=CLng(Alpha(13)), _
                       lngPenStyle:=zwStyleSolid, lngPenWidth:=zwWidthThin
    .DataBase.Add Prim:=rpdPolyData
    Set rpdPolyData = Nothing
    Set rpdDPoints = Nothing
    
    rpdLineData.Start.x = dblOutCX + (dblOutML / 2) * (i * 2 - 3) + _
                                     (i * 15.5 - 20.5) / .DrawInfo.Numerator
    rpdLineData.Start.y = dblOutCY
    rpdLineData.End.x = dblOutCX + dblOutML / 2 + _
                                   (i * 15.5 - 6) / .DrawInfo.Numerator
    rpdLineData.End.y = dblOutCY
    objRapidPrim.SetUp prmObject:=rpdLineData, lngLayerNo:=CLng(Alpha(11)), _
                       lngPenStyle:=zwStyleS1DashDot, lngPenWidth:=zwWidthThin
    .DataBase.Add Prim:=rpdLineData
  Next
End With
Set rpdLineData = Nothing
Set rpdDPoint1 = Nothing
Set rpdDPoint2 = Nothing
Set rpdDPoint3 = Nothing
Set objRapidPrim = Nothing
Set objRapidLine = Nothing
Set objGPData = Nothing

End Sub


Public Function varElement(ByVal strEName As String, ByVal strEParam As String) As Variant

Dim i As Long, j As Long

varElement = -1
Select Case strEParam
  Case "POS": j = 2
  Case "DIA": j = 3
  Case "EF1": j = 4
  Case "EF2": j = 5
  Case "PPX": j = 6
  Case "PPY": j = 7
  Case "CT1": j = 8
  Case "CT2": j = 9
  Case "BAL": j = 10
  Case Else: Exit Function
End Select
For i = 1 To lngOutMC
  If varOutE(i, 1) = strEName Then
    varElement = varOutE(i, j)
    Exit For
  End If
Next

End Function


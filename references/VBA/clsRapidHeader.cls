VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsRapidHeader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False


Option Explicit
Option Base 1
Option Compare Binary


Public Sub AddA(ByRef apdObject As clsAssyPlotData)

Dim rpdLineData As zwDrawCAD.LineData
Dim rpdArcData As zwDrawCAD.ArcData
Dim rpdTextData As zwDrawCAD.TextData
Dim objRapidPrim As clsRapidPrim
Dim objRapidParam As clsRapidParam

Dim i As Long, j As Long, n As Long
Dim rc As Long, dd As Long
Dim maxRow As Long
Dim ff(4, 6) As Long, ll(99) As Long
Dim xx As Double, yy As Double
Dim hh As Double, ww As Double
Dim tmpx As Double, tmpy As Double
Dim sLine As Double, eLine As Double
Dim strFont As String, strGName(99) As String

ff(1, 1) = 1: ff(1, 2) = -1: ff(1, 3) = 1: ff(1, 4) = 0: ff(1, 5) = 0: ff(1, 6) = -1
ff(2, 1) = -1: ff(2, 2) = -1: ff(2, 3) = 0: ff(2, 4) = -1: ff(2, 5) = -1: ff(2, 6) = 0
ff(3, 1) = -1: ff(3, 2) = 1: ff(3, 3) = -1: ff(3, 4) = 0: ff(3, 5) = 0: ff(3, 6) = 1
ff(4, 1) = 1: ff(4, 2) = 1: ff(4, 3) = 0: ff(4, 4) = 1: ff(4, 5) = 1: ff(4, 6) = 0

Set objRapidParam = New clsRapidParam
objRapidParam.Import strConfig:="Param"
With objRapidParam
  strFont = .strItem(1)
  
  rc = 0
  For i = 1 To .lngMaxCount
    If .strItem(i) = "è≈ì_ãóó£ìô" Then
      yy = .dblXY(i, 2) / rpdDoc.DrawInfo.Numerator
      hh = .dblHW(i, 1) / rpdDoc.DrawInfo.Numerator
      ww = .dblHW(i, 2) / rpdDoc.DrawInfo.Numerator
      dd = .lngVal(i, 1)
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
    lngGla = 0: xx = .dblCenterX
    For i = 1 To .lngMaxCount
      If Left(.strEName(i), 1) = "G" Then
        lngGla = lngGla + 1
        strGName(lngGla) = .strEName(i)
        If .varElement(.strEName(i), "EF1") > _
           .varElement(.strEName(i), "EF2") Then
          ll(lngGla) = Len(Fix(.varElement(.strEName(i), "EF1"))) + dd + 3
        Else
          ll(lngGla) = Len(Fix(.varElement(.strEName(i), "EF2"))) + dd + 3
        End If
        xx = xx - (ll(lngGla) + 2) * (ww / 4)
      End If
      
      maxRow = Cells(7, 2).End(xlDown).Row
      If maxRow > 7 Then
        For j = 8 To maxRow
          If Not IsEmpty(Cells(j, 3).Value) Then
            If Right(Cells(j, 2).Value, 3) = Left(.strEName(i), 3) Then
              lngGla = lngGla + 1
              strGName(lngGla) = Cells(j, 3).Text
              xx = xx - ww / 4
              Exit For
            End If
          End If
        Next
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
Set objRapidPrim = New clsRapidPrim
With apdObject
  xx = xx + ww / 4
  sLine = xx
  For i = 1 To lngGla
    objProgCtrl.AddBar lngDivCount:=lngGla
    If Left(strGName(i), 2) <> "f=" Then
      xx = xx + (ll(i) + 2) * (ww / 4)
      
      rpdLineData.Start.x = xx - ll(i) * (ww / 4)
      rpdLineData.Start.y = yy
      rpdLineData.End.x = xx + ll(i) * (ww / 4)
      rpdLineData.End.y = yy
      objRapidPrim.SetUp prmObject:=rpdLineData, lngLayerNo:=CLng(Alpha(12)), _
                         lngPenStyle:=zwStyleSolid, lngPenWidth:=zwWidthThin
      rpdDoc.DataBase.Add Prim:=rpdLineData
      
      For n = 1 To 6
        Set rpdTextData = New zwDrawCAD.TextData
        rpdTextData.FaceName = strFont
        rpdTextData.Height = hh
        rpdTextData.Width = ww
        rpdTextData.BasePosition = 4
        rpdTextData.BasePoint.x = xx
        rpdTextData.BasePoint.y = yy
        Select Case n
          Case 1
            rpdTextData.Text = Format(.varElement(strGName(i), "DIA"), _
                                      """É”""0." & String(dd, "0"))
            rpdTextData.BasePoint.y = yy + hh * 2.7
          Case 2
            rpdTextData.Text = Format(.varElement(strGName(i), "EF1"), _
                                      """É”""0." & String(dd, "0"))
            rpdTextData.BasePoint.y = yy + hh * 0.6
          Case 3
            rpdTextData.Text = Format(.varElement(strGName(i), "EF2"), _
                                      """É”""0." & String(dd, "0"))
            rpdTextData.BasePoint.y = yy - hh * 0.6
          Case 4
            rpdTextData.Text = .varElement(strGName(i), "CT1")
            rpdTextData.BasePosition = 2
            rpdTextData.BasePoint.x = xx - (ll(i) - 3) * (ww / 4)
            rpdTextData.BasePoint.y = yy + hh * 1.1
          Case 5
            rpdTextData.Text = .varElement(strGName(i), "CT2")
            rpdTextData.BasePosition = 6
            rpdTextData.BasePoint.x = xx + (ll(i) - 3) * (ww / 4)
            rpdTextData.BasePoint.y = yy - hh * 1.1
          Case 6
            rpdTextData.Text = "(" & Space(ll(i)) & ")"
            rpdTextData.Height = hh * 2.2
        End Select
        objRapidPrim.SetUp prmObject:=rpdTextData, lngLayerNo:=CLng(Alpha(12)), _
                           lngPenStyle:=zwStyleSolid, lngPenWidth:=zwWidthThin
        rpdDoc.DataBase.Add Prim:=rpdTextData
        Set rpdTextData = Nothing
      Next
      
      xx = xx + (ll(i) + 2) * (ww / 4)
    Else
      eLine = xx
      
      For n = 1 To 4
        Select Case n
          Case 1: tmpx = sLine
          Case 2: tmpx = eLine
          Case Else: tmpx = (sLine + eLine) / 2
        End Select
        tmpy = yy + hh * 3.7
        
        Set rpdArcData = New zwDrawCAD.ArcData
        rpdArcData.Radius = hh / 2
        rpdArcData.Center.x = tmpx + (hh / 2) * ff(n, 1)
        rpdArcData.Center.y = tmpy + (hh / 2) * ff(n, 2)
        rpdArcData.Start.x = tmpx + (hh / 2) * ff(n, 3)
        rpdArcData.Start.y = tmpy + (hh / 2) * ff(n, 4)
        rpdArcData.End.x = tmpx + (hh / 2) * ff(n, 5)
        rpdArcData.End.y = tmpy + (hh / 2) * ff(n, 6)
        objRapidPrim.SetUp prmObject:=rpdArcData, lngLayerNo:=CLng(Alpha(12)), _
                           lngPenStyle:=zwStyleSolid, lngPenWidth:=zwWidthThin
        rpdDoc.DataBase.Add Prim:=rpdArcData
        Set rpdArcData = Nothing
      Next
      
      For n = 1 To 2
        Select Case n
          Case 1
            rpdLineData.Start.x = sLine + hh / 2
            rpdLineData.End.x = (sLine + eLine) / 2 - hh / 2
          Case 2
            rpdLineData.Start.x = eLine - hh / 2
            rpdLineData.End.x = (sLine + eLine) / 2 + hh / 2
        End Select
        rpdLineData.Start.y = yy + hh * 3.7
        rpdLineData.End.y = yy + hh * 3.7
        objRapidPrim.SetUp prmObject:=rpdLineData, lngLayerNo:=CLng(Alpha(12)), _
                           lngPenStyle:=zwStyleSolid, lngPenWidth:=zwWidthThin
        rpdDoc.DataBase.Add Prim:=rpdLineData
      Next
      
      Set rpdTextData = New zwDrawCAD.TextData
      rpdTextData.Text = strGName(i)
      rpdTextData.FaceName = strFont
      rpdTextData.Height = hh
      rpdTextData.Width = ww
      rpdTextData.BasePosition = 4
      rpdTextData.BasePoint.x = (sLine + eLine) / 2
      rpdTextData.BasePoint.y = yy + hh * 4.8
      objRapidPrim.SetUp prmObject:=rpdTextData, lngLayerNo:=CLng(Alpha(12)), _
                         lngPenStyle:=zwStyleSolid, lngPenWidth:=zwWidthThin
      rpdDoc.DataBase.Add Prim:=rpdTextData
      Set rpdTextData = Nothing
      
      xx = xx + ww / 2
      sLine = xx
    End If
  Next
End With
Set rpdLineData = Nothing
Set objRapidPrim = Nothing

End Sub


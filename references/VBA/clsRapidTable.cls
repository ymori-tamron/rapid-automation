VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsRapidTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False


Option Explicit
Option Base 1
Option Compare Binary


Public Sub AddA()

Dim rpdLineData As zwDrawCAD.LineData
Dim rpdTextData As zwDrawCAD.TextData
Dim objRapidPrim As clsRapidPrim
Dim objRapidParam As clsRapidParam

Dim mrg As Boolean
Dim i As Long, j As Long, n As Long, rc As Long
Dim maxRow As Long, maxCol As Long, maxChr As Long
Dim xx As Double, yy As Double, zz As Double
Dim hh As Double, ww As Double, tt(2) As Double
Dim strFont As String, txt(2) As String

Set objRapidParam = New clsRapidParam
objRapidParam.Import strConfig:="Param"
With objRapidParam
  strFont = .strItem(1)
  
  rc = 0
  For i = 1 To .lngMaxCount
    If .strItem(i) = "光学仕様書" Then
      xx = .dblXY(i, 1) / rpdDoc.DrawInfo.Numerator
      yy = .dblXY(i, 2) / rpdDoc.DrawInfo.Numerator
      hh = .dblHW(i, 1) / rpdDoc.DrawInfo.Numerator
      ww = .dblHW(i, 2) / rpdDoc.DrawInfo.Numerator
      rc = i: Exit For
    End If
  Next
End With
Set objRapidParam = Nothing

If rc = 0 Then
  objProgCtrl.AddBar lngDivCount:=1
  Exit Sub
End If

maxRow = Cells(23, 5).End(xlDown).Row
If maxRow > 23 Then
  maxCol = 8
  For j = 6 To 8
    If IsEmpty(Cells(24, j).Value) Then
      maxCol = j - 1
      Exit For
    End If
  Next
  
  If maxCol < 6 Then
    objProgCtrl.AddBar lngDivCount:=1
    Exit Sub
  End If
Else
  objProgCtrl.AddBar lngDivCount:=1
  Exit Sub
End If

Set rpdLineData = New zwDrawCAD.LineData
Set objRapidPrim = New clsRapidPrim
With rpdLineData
  .Start.x = xx - ww * 6
  .Start.y = yy - hh * 1.2
  .End.x = xx + ww * 6
  .End.y = yy - hh * 1.2
  objRapidPrim.SetUp prmObject:=rpdLineData, lngLayerNo:=CLng(Alpha(12)), _
                     lngPenStyle:=zwStyleSolid, lngPenWidth:=zwWidthThin
  rpdDoc.DataBase.Add Prim:=rpdLineData
End With

Set rpdTextData = New zwDrawCAD.TextData
With rpdTextData
  .Text = "光学仕様書"
  .FaceName = strFont
  .Height = hh * 2
  .Width = ww * 2
  .BasePosition = 4
  .BasePoint.x = xx
  .BasePoint.y = yy
  objRapidPrim.SetUp prmObject:=rpdTextData, lngLayerNo:=CLng(Alpha(12)), _
                     lngPenStyle:=zwStyleSolid, lngPenWidth:=zwWidthThin
  rpdDoc.DataBase.Add Prim:=rpdTextData
End With
Set rpdTextData = Nothing

mrg = False
txt(1) = "": txt(2) = ""
xx = xx - ww * 16
yy = yy - hh * 2
zz = 0
If maxCol = 8 Then
  tt(1) = ww * 11
  tt(2) = ww * 7
Else
  tt(1) = (ww * 32) / (maxCol - 4)
  tt(2) = (ww * 32) / (maxCol - 4)
End If
For i = 24 To maxRow
  objProgCtrl.AddBar lngDivCount:=(maxRow - 23)
  If InStr(Cells(i, 5).Value, "/") > 0 Then
    txt(1) = Left(Cells(i, 5).Value, InStr(Cells(i, 5).Value, "/") - 1)
    txt(2) = Mid(Cells(i, 5).Value, InStr(Cells(i, 5).Value, "/") + 1)
    zz = tt(1) - (LenB(StrConv(txt(2), vbFromUnicode)) + 2) * (ww / 3)
  Else
    txt(1) = Cells(i, 5).Value
    txt(2) = ""
    zz = 0
  End If
  
  With rpdLineData
    For n = 1 To 2
      .Start.x = xx
      .Start.y = yy - (hh * 2) * (i - 24)
      Select Case n
        Case 1
          If mrg Then .Start.x = xx + zz
          .End.x = xx + ww * 32
          .End.y = yy - (hh * 2) * (i - 24)
        Case 2
          .End.x = xx
          .End.y = yy - (hh * 2) * (i - 23)
      End Select
      objRapidPrim.SetUp prmObject:=rpdLineData, lngLayerNo:=CLng(Alpha(12)), _
                         lngPenStyle:=zwStyleSolid, lngPenWidth:=zwWidthThin
      rpdDoc.DataBase.Add Prim:=rpdLineData
    Next
    
    If zz > 0 Then
      .Start.x = xx + zz
      .Start.y = yy - (hh * 2) * (i - 24)
      .End.x = xx + zz
      .End.y = yy - (hh * 2) * (i - 23)
      objRapidPrim.SetUp prmObject:=rpdLineData, lngLayerNo:=CLng(Alpha(12)), _
                         lngPenStyle:=zwStyleSolid, lngPenWidth:=zwWidthThin
      rpdDoc.DataBase.Add Prim:=rpdLineData
    End If
    
    .Start.x = xx + tt(1)
    .Start.y = yy - (hh * 2) * (i - 24)
    .End.x = xx + tt(1)
    .End.y = yy - (hh * 2) * (i - 23)
    objRapidPrim.SetUp prmObject:=rpdLineData, lngLayerNo:=CLng(Alpha(12)), _
                       lngPenStyle:=zwStyleSolid, lngPenWidth:=zwWidthThin
    rpdDoc.DataBase.Add Prim:=rpdLineData
  End With
  
  If Not mrg Then
    Set rpdTextData = New zwDrawCAD.TextData
    With rpdTextData
      .Text = txt(1)
      .FaceName = strFont
      .Height = hh
      .Width = ww
      If zz > 0 Then
        maxChr = zz / (ww / 2) - 2
      Else
        maxChr = tt(1) / (ww / 2) - 2
      End If
      If LenB(StrConv(txt(1), vbFromUnicode)) > maxChr Then
        .Width = ww * (maxChr / LenB(StrConv(txt(1), vbFromUnicode)))
      End If
      .BasePosition = 4
      If zz > 0 Then
        .BasePoint.x = xx + zz / 2
        .BasePoint.y = yy - (hh * 2) * (i - 23)
      Else
        .BasePoint.x = xx + tt(1) / 2
        .BasePoint.y = yy - (hh * 2) * (i - 23.5)
        If Cells(i, 5).Value = "最短撮影距離(MOD)" Or _
           Cells(i, 5).Value = "基準テスト位置" Then
          .BasePoint.y = yy - (hh * 2) * (i - 23.63)
        End If
      End If
      objRapidPrim.SetUp prmObject:=rpdTextData, lngLayerNo:=CLng(Alpha(12)), _
                         lngPenStyle:=zwStyleSolid, lngPenWidth:=zwWidthThin
      rpdDoc.DataBase.Add Prim:=rpdTextData
    End With
    Set rpdTextData = Nothing
  End If
  
  If zz > 0 Then
    Set rpdTextData = New zwDrawCAD.TextData
    With rpdTextData
      .Text = txt(2)
      .FaceName = strFont
      .Height = hh / 1.5
      .Width = ww / 1.5
      .BasePosition = 4
      .BasePoint.x = xx + tt(1) / 2 + zz / 2
      .BasePoint.y = yy - (hh * 2) * (i - 23.5)
      objRapidPrim.SetUp prmObject:=rpdTextData, lngLayerNo:=CLng(Alpha(12)), _
                         lngPenStyle:=zwStyleSolid, lngPenWidth:=zwWidthThin
      rpdDoc.DataBase.Add Prim:=rpdTextData
    End With
    Set rpdTextData = Nothing
  Else
    If Cells(i, 5).Value = "最短撮影距離(MOD)" Or _
       Cells(i, 5).Value = "基準テスト位置" Then
      Set rpdTextData = New zwDrawCAD.TextData
      With rpdTextData
        .Text = "(" & Trim(Mid(Cells(23, 8).Value, 8)) & ")"
        .FaceName = strFont
        .Height = hh / 2
        .Width = ww / 2
        .BasePosition = 4
        .BasePoint.x = xx + tt(1) / 2
        .BasePoint.y = yy - (hh * 2) * (i - 23.17)
        objRapidPrim.SetUp prmObject:=rpdTextData, lngLayerNo:=CLng(Alpha(12)), _
                           lngPenStyle:=zwStyleSolid, lngPenWidth:=zwWidthThin
        rpdDoc.DataBase.Add Prim:=rpdTextData
      End With
      Set rpdTextData = Nothing
    End If
  End If
  
  For j = 6 To maxCol
    With rpdLineData
      .Start.x = xx + tt(1) + tt(2) * (j - 5)
      .Start.y = yy - (hh * 2) * (i - 24)
      .End.x = xx + tt(1) + tt(2) * (j - 5)
      .End.y = yy - (hh * 2) * (i - 23)
      objRapidPrim.SetUp prmObject:=rpdLineData, lngLayerNo:=CLng(Alpha(12)), _
                         lngPenStyle:=zwStyleSolid, lngPenWidth:=zwWidthThin
      rpdDoc.DataBase.Add Prim:=rpdLineData
    End With
    
    Set rpdTextData = New zwDrawCAD.TextData
    With rpdTextData
      .Text = Trim(Cells(i, j).Text)
      .FaceName = strFont
      .Height = hh
      .Width = ww
      maxChr = tt(2) / (ww / 2) - 2
      If LenB(StrConv(Trim(Cells(i, j).Text), vbFromUnicode)) > maxChr Then
        .Width = ww * (maxChr / LenB(StrConv(Trim(Cells(i, j).Text), vbFromUnicode)))
      End If
      .BasePosition = 4
      .BasePoint.x = xx + tt(1) + tt(2) * (j - 5.5)
      .BasePoint.y = yy - (hh * 2) * (i - 23.5)
      objRapidPrim.SetUp prmObject:=rpdTextData, lngLayerNo:=CLng(Alpha(12)), _
                         lngPenStyle:=zwStyleSolid, lngPenWidth:=zwWidthThin
      rpdDoc.DataBase.Add Prim:=rpdTextData
    End With
    Set rpdTextData = Nothing
  Next
  
  If InStr(Cells(i, 5).Value, "/") > 0 Then
    Select Case mrg
      Case True: mrg = False
      Case False: mrg = True
    End Select
  End If
Next

With rpdLineData
  .Start.x = xx
  .Start.y = yy - (hh * 2) * (maxRow - 23)
  .End.x = xx + ww * 32
  .End.y = yy - (hh * 2) * (maxRow - 23)
  objRapidPrim.SetUp prmObject:=rpdLineData, lngLayerNo:=CLng(Alpha(12)), _
                     lngPenStyle:=zwStyleSolid, lngPenWidth:=zwWidthThin
  rpdDoc.DataBase.Add Prim:=rpdLineData
End With
Set rpdLineData = Nothing
Set objRapidPrim = Nothing

End Sub


VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsRapidTol"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False


Option Explicit
Option Base 1
Option Compare Binary

Private strOut As String


Private Sub Class_Initialize()

strOut = ""

End Sub


Public Property Get strKeyWord() As String

strKeyWord = strOut

End Property


Public Sub AddA(ByRef rpObject As clsRapidParam, _
                ByVal lngPRow As Long, ByVal lngTRow As Long, _
                ByVal dblBaseX As Double, ByVal dblBaseY As Double, _
                ByVal dblShift As Double, ByVal dblAngle As Double)

Dim rpdTextData As zwDrawCAD.TextData
Dim objRapidPrim As clsRapidPrim

Dim fmt As String, txt(2) As String

If lngTRow > 0 Then
  If InStr(Cells(lngTRow, 3).Value, "/") > 0 Then
    If IsNumeric(Left(Cells(lngTRow, 3).Value, InStr(Cells(lngTRow, 3).Value, "/") - 1)) And _
       IsNumeric(Mid(Cells(lngTRow, 3).Value, InStr(Cells(lngTRow, 3).Value, "/") + 1)) Then
      txt(1) = Left(Cells(lngTRow, 3).Value, InStr(Cells(lngTRow, 3).Value, "/") - 1)
      txt(2) = Mid(Cells(lngTRow, 3).Value, InStr(Cells(lngTRow, 3).Value, "/") + 1)
    Else
      Exit Sub
    End If
  Else
    Exit Sub
  End If
Else
  Exit Sub
End If

Set objRapidPrim = New clsRapidPrim
If CDbl(txt(1)) + CDbl(txt(2)) = 0 Then
  fmt = """Å}""0.0##"
Else
  fmt = "+0.0##;-0.0##"
End If
Set rpdTextData = New zwDrawCAD.TextData
rpdTextData.Text = Format(CDbl(txt(1)), fmt)
rpdTextData.FaceName = rpObject.strItem(1)
rpdTextData.Height = rpObject.dblHW(lngPRow, 1) / rpdDoc.DrawInfo.Numerator
rpdTextData.Width = rpObject.dblHW(lngPRow, 2) / rpdDoc.DrawInfo.Numerator
rpdTextData.BasePosition = 0
Select Case dblAngle
  Case 0
    rpdTextData.BasePoint.x = dblBaseX + dblShift
    rpdTextData.BasePoint.y = dblBaseY + _
      (rpObject.dblHW(lngPRow, 1) * 1.2) / rpdDoc.DrawInfo.Numerator
  Case 90
    rpdTextData.BasePoint.x = dblBaseX - _
      (rpObject.dblHW(lngPRow, 1) * 1.2) / rpdDoc.DrawInfo.Numerator
    rpdTextData.BasePoint.y = dblBaseY + dblShift
  Case -90, 270
    rpdTextData.BasePoint.x = dblBaseX + _
      (rpObject.dblHW(lngPRow, 1) * 1.2) / rpdDoc.DrawInfo.Numerator
    rpdTextData.BasePoint.y = dblBaseY - dblShift
  Case Else
    rpdTextData.BasePoint.x = dblBaseX
    rpdTextData.BasePoint.y = dblBaseY
End Select
rpdTextData.Angle = dblAngle
objRapidPrim.SetUp prmObject:=rpdTextData, lngLayerNo:=CLng(Alpha(13)), _
                   lngPenStyle:=zwStyleSolid, lngPenWidth:=zwWidthThin
rpdDoc.DataBase.Add Prim:=rpdTextData
Set rpdTextData = Nothing

If CDbl(txt(1)) + CDbl(txt(2)) <> 0 Then
  Set rpdTextData = New zwDrawCAD.TextData
  rpdTextData.Text = Format(CDbl(txt(2)), fmt)
  rpdTextData.FaceName = rpObject.strItem(1)
  rpdTextData.Height = rpObject.dblHW(lngPRow, 1) / rpdDoc.DrawInfo.Numerator
  rpdTextData.Width = rpObject.dblHW(lngPRow, 2) / rpdDoc.DrawInfo.Numerator
  rpdTextData.BasePosition = 0
  Select Case dblAngle
    Case 0
      rpdTextData.BasePoint.x = dblBaseX + dblShift
      rpdTextData.BasePoint.y = dblBaseY
    Case 90
      rpdTextData.BasePoint.x = dblBaseX
      rpdTextData.BasePoint.y = dblBaseY + dblShift
    Case -90, 270
      rpdTextData.BasePoint.x = dblBaseX
      rpdTextData.BasePoint.y = dblBaseY - dblShift
    Case Else
      rpdTextData.BasePoint.x = dblBaseX
      rpdTextData.BasePoint.y = dblBaseY
  End Select
  rpdTextData.Angle = dblAngle
  objRapidPrim.SetUp prmObject:=rpdTextData, lngLayerNo:=CLng(Alpha(13)), _
                     lngPenStyle:=zwStyleSolid, lngPenWidth:=zwWidthThin
  rpdDoc.DataBase.Add Prim:=rpdTextData
  Set rpdTextData = Nothing
End If
Set objRapidPrim = Nothing

End Sub


Public Sub AddG(ByRef rpObject As clsRapidParam, ByVal lngRow As Long, _
                ByVal dblBaseX As Double, ByVal dblBaseY As Double, _
                ByVal dblShift As Double, ByVal dblAngle As Double)

Dim rpdTextData As zwDrawCAD.TextData
Dim objRapidPrim As clsRapidPrim

Dim i As Long, cc As Long, maxCol As Long
Dim txt(2) As String

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
    If Not IsEmpty(.Cells(rpObject.lngVal(lngRow, 1), cc).Value) And _
       Not IsEmpty(.Cells(rpObject.lngVal(lngRow, 1), cc + 1).Value) Then
      If IsNumeric(.Cells(rpObject.lngVal(lngRow, 1), cc).Value) And _
         IsNumeric(.Cells(rpObject.lngVal(lngRow, 1), cc + 1).Value) Then
        txt(1) = .Cells(rpObject.lngVal(lngRow, 1), cc).Text
        txt(2) = .Cells(rpObject.lngVal(lngRow, 1), cc + 1).Text
      Else
        Exit Sub
      End If
    Else
      Exit Sub
    End If
  Else
    Exit Sub
  End If
  
  Set objRapidPrim = New clsRapidPrim
  If .Cells(rpObject.lngVal(lngRow, 1), cc).Value + _
     .Cells(rpObject.lngVal(lngRow, 1), cc + 1).Value = 0 Then
    txt(1) = "Å}" & Mid(txt(1), 2)
  End If
  Set rpdTextData = New zwDrawCAD.TextData
  rpdTextData.Text = txt(1)
  rpdTextData.FaceName = rpObject.strItem(1)
  rpdTextData.Height = rpObject.dblHW(lngRow, 1) / rpdDoc.DrawInfo.Numerator
  rpdTextData.Width = rpObject.dblHW(lngRow, 2) / rpdDoc.DrawInfo.Numerator
  rpdTextData.BasePosition = 0
  Select Case dblAngle
    Case 0
      rpdTextData.BasePoint.x = dblBaseX + dblShift
      rpdTextData.BasePoint.y = dblBaseY + _
        (rpObject.dblHW(lngRow, 1) * 1.2) / rpdDoc.DrawInfo.Numerator
    Case 90
      rpdTextData.BasePoint.x = dblBaseX - _
        (rpObject.dblHW(lngRow, 1) * 1.2) / rpdDoc.DrawInfo.Numerator
      rpdTextData.BasePoint.y = dblBaseY + dblShift
    Case -90, 270
      rpdTextData.BasePoint.x = dblBaseX + _
        (rpObject.dblHW(lngRow, 1) * 1.2) / rpdDoc.DrawInfo.Numerator
      rpdTextData.BasePoint.y = dblBaseY - dblShift
    Case Else
      rpdTextData.BasePoint.x = dblBaseX
      rpdTextData.BasePoint.y = dblBaseY
  End Select
  rpdTextData.Angle = dblAngle
  objRapidPrim.SetUp prmObject:=rpdTextData, lngLayerNo:=CLng(Alpha(13)), _
                     lngPenStyle:=zwStyleSolid, lngPenWidth:=zwWidthThin
  rpdDoc.DataBase.Add Prim:=rpdTextData
  Set rpdTextData = Nothing
  
  If .Cells(rpObject.lngVal(lngRow, 1), cc).Value + _
     .Cells(rpObject.lngVal(lngRow, 1), cc + 1).Value <> 0 Then
    Set rpdTextData = New zwDrawCAD.TextData
    rpdTextData.Text = txt(2)
    rpdTextData.FaceName = rpObject.strItem(1)
    rpdTextData.Height = rpObject.dblHW(lngRow, 1) / rpdDoc.DrawInfo.Numerator
    rpdTextData.Width = rpObject.dblHW(lngRow, 2) / rpdDoc.DrawInfo.Numerator
    rpdTextData.BasePosition = 0
    Select Case dblAngle
      Case 0
        rpdTextData.BasePoint.x = dblBaseX + dblShift
        rpdTextData.BasePoint.y = dblBaseY
      Case 90
        rpdTextData.BasePoint.x = dblBaseX
        rpdTextData.BasePoint.y = dblBaseY + dblShift
      Case -90, 270
        rpdTextData.BasePoint.x = dblBaseX
        rpdTextData.BasePoint.y = dblBaseY - dblShift
      Case Else
        rpdTextData.BasePoint.x = dblBaseX
        rpdTextData.BasePoint.y = dblBaseY
    End Select
    rpdTextData.Angle = dblAngle
    objRapidPrim.SetUp prmObject:=rpdTextData, lngLayerNo:=CLng(Alpha(13)), _
                       lngPenStyle:=zwStyleSolid, lngPenWidth:=zwWidthThin
    rpdDoc.DataBase.Add Prim:=rpdTextData
    Set rpdTextData = Nothing
  End If
  Set objRapidPrim = Nothing
End With

End Sub


Public Sub AddP(ByRef rpObject As clsRapidParam, ByVal lngRow As Long, _
                ByVal dblBaseX As Double, ByVal dblBaseY As Double, _
                ByVal dblShift As Double, ByVal dblAngle As Double, _
                ByVal strType As String, ByVal lngZD As Long)

Dim rpdTextData As zwDrawCAD.TextData
Dim objRapidPrim As clsRapidPrim
Dim objPressParam As clsPressParam

Dim i As Long
Dim txt(2) As String

Set objPressParam = New clsPressParam
objPressParam.Import strConfig:="PressParam"
With objPressParam
  txt(1) = "+0. ": txt(2) = ""
  Select Case strType
    Case "DP"
      For i = 1 To .lngMaxCount("DPåˆç∑")
        If Cells(8, 6).Value < .dbltDP(i, 9) Then
          Select Case rpObject.strItem(lngRow)
            Case "ì˜å˙åˆç∑": txt(1) = "Å}" & Format(.dbltDP(i, lngZD), "0.0#")
            Case "äOåaåˆç∑": txt(2) = "-" & Format(.dbltDP(i, lngZD), "0.0#")
          End Select
          Exit For
        End If
      Next
      
    Case "HP"
      For i = 1 To .lngMaxCount("HPåˆç∑")
        If Cells(8, 6).Value < .dbltHP(i, 9) Then
          Select Case rpObject.strItem(lngRow)
            Case "ì˜å˙åˆç∑": txt(1) = "Å}" & Format(.dbltHP(i, lngZD), "0.0#")
            Case "äOåaåˆç∑": txt(2) = "-" & Format(.dbltHP(i, lngZD), "0.0#")
          End Select
          Exit For
        End If
      Next
  End Select
End With
Set objPressParam = Nothing

Set objRapidPrim = New clsRapidPrim
Set rpdTextData = New zwDrawCAD.TextData
With rpdTextData
  .Text = txt(1)
  .FaceName = rpObject.strItem(1)
  .Height = rpObject.dblHW(lngRow, 1) / rpdDoc.DrawInfo.Numerator
  .Width = rpObject.dblHW(lngRow, 2) / rpdDoc.DrawInfo.Numerator
  .BasePosition = 0
  Select Case dblAngle
    Case 0
      .BasePoint.x = dblBaseX + dblShift
      .BasePoint.y = dblBaseY + _
        (rpObject.dblHW(lngRow, 1) * 1.2) / rpdDoc.DrawInfo.Numerator
    Case 90
      .BasePoint.x = dblBaseX - _
        (rpObject.dblHW(lngRow, 1) * 1.2) / rpdDoc.DrawInfo.Numerator
      .BasePoint.y = dblBaseY + dblShift
    Case -90, 270
      .BasePoint.x = dblBaseX + _
        (rpObject.dblHW(lngRow, 1) * 1.2) / rpdDoc.DrawInfo.Numerator
      .BasePoint.y = dblBaseY - dblShift
    Case Else
      .BasePoint.x = dblBaseX
      .BasePoint.y = dblBaseY
  End Select
  .Angle = dblAngle
  objRapidPrim.SetUp prmObject:=rpdTextData, lngLayerNo:=lngPre(4), _
                     lngPenStyle:=zwStyleSolid, lngPenWidth:=zwWidthThin
  rpdDoc.DataBase.Add Prim:=rpdTextData
End With
Set rpdTextData = Nothing

If txt(2) <> "" Then
  Set rpdTextData = New zwDrawCAD.TextData
  With rpdTextData
    .Text = txt(2)
    .FaceName = rpObject.strItem(1)
    .Height = rpObject.dblHW(lngRow, 1) / rpdDoc.DrawInfo.Numerator
    .Width = rpObject.dblHW(lngRow, 2) / rpdDoc.DrawInfo.Numerator
    .BasePosition = 0
    Select Case dblAngle
      Case 0
        .BasePoint.x = dblBaseX + dblShift
        .BasePoint.y = dblBaseY
      Case 90
        .BasePoint.x = dblBaseX
        .BasePoint.y = dblBaseY + dblShift
      Case -90, 270
        .BasePoint.x = dblBaseX
        .BasePoint.y = dblBaseY - dblShift
      Case Else
        .BasePoint.x = dblBaseX
        .BasePoint.y = dblBaseY
    End Select
    .Angle = dblAngle
    objRapidPrim.SetUp prmObject:=rpdTextData, lngLayerNo:=lngPre(4), _
                       lngPenStyle:=zwStyleSolid, lngPenWidth:=zwWidthThin
    rpdDoc.DataBase.Add Prim:=rpdTextData
  End With
  Set rpdTextData = Nothing
End If
Set objRapidPrim = Nothing

End Sub


Public Sub AddB(ByRef rpObject As clsRapidParam, ByVal lngRow As Long, _
                ByVal dblBaseX As Double, ByVal dblBaseY As Double, _
                ByVal dblShift As Double, ByVal dblAngle As Double)

Dim rpdTextData As zwDrawCAD.TextData
Dim objRapidPrim As clsRapidPrim

Dim chk As Boolean
Dim i As Long, j As Long, m As Long
Dim rc As Long, cc As Long, maxCol As Long
Dim tmp(2) As Double
Dim fmt As String, prtName(4) As String

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

With Worksheets(DefaultSheet3)
  tmp(1) = 0: tmp(2) = 0
  If Left(ActiveSheet.Name, 1) = "B" Then
    If InStr(prtName(1), "ASP") = 0 And _
       Worksheets(prtName(1)).Cells(14, 9).Value > 0 Then
      tmp(1) = 0.01: tmp(2) = 0.01
    End If
    If InStr(prtName(2), "ASP") = 0 And _
       Worksheets(prtName(2)).Cells(14, 9).Value > 0 Then
      tmp(1) = 0.01: tmp(2) = 0.01
    End If
  End If
  
  For j = 1 To 4
    Select Case j
      Case 1, 2
        chk = True
      Case 3, 4
        chk = False
        If Left(ActiveSheet.Name, 1) = "B" Then
          If Left(Cells(j + 5, 3).Value, 1) = "M" Then
            prtName(j) = "H" & Mid(Cells(j + 5, 3).Value, 2, 2) & "(ASP)"
            chk = True
          End If
        End If
    End Select
    
    If chk Then
      rc = 0
      If Left(ActiveSheet.Name, 1) = "M" And _
         rpObject.strItem(lngRow) = "áôÇgåˆç∑" & CStr(j) Then
        For i = 1 To rpObject.lngMaxCount
          If rpObject.strItem(i) = "ì˜å˙åˆç∑" Then
            rc = i: Exit For
          End If
        Next
      Else
        rc = lngRow
      End If
      
      cc = 0
      maxCol = .Cells(6, .Columns.Count).End(xlToLeft).Column
      If maxCol > 2 Then
        For i = 3 To maxCol Step 3
          If .Cells(6, i).Value = prtName(j) Then
            cc = i: Exit For
          End If
        Next
      End If
      
      If rc > 0 And _
         cc > 0 Then
        For m = 1 To 2
          If Not IsEmpty(.Cells(rpObject.lngVal(rc, 1), cc + (m - 1)).Value) And _
             IsNumeric(.Cells(rpObject.lngVal(rc, 1), cc + (m - 1)).Value) Then
            tmp(m) = tmp(m) + .Cells(rpObject.lngVal(rc, 1), cc + (m - 1)).Value
          Else
            tmp(m) = tmp(m) + 0.1 * (m * (-2) + 3)
          End If
        Next
      Else
        tmp(1) = tmp(1) + 0.1: tmp(2) = tmp(2) - 0.1
      End If
    End If
  Next
  
  If Left(ActiveSheet.Name, 1) = "B" Then
    If Cells(15, 3).Value <> "+[ó›êœ]/-[ó›êœ]" And _
       InStr(Cells(15, 3).Value, "/") > 0 Then
      If IsNumeric(Left(Cells(15, 3).Value, InStr(Cells(15, 3).Value, "/") - 1)) And _
         IsNumeric(Mid(Cells(15, 3).Value, InStr(Cells(15, 3).Value, "/") + 1)) Then
        tmp(1) = Left(Cells(15, 3).Value, InStr(Cells(15, 3).Value, "/") - 1)
        tmp(2) = Mid(Cells(15, 3).Value, InStr(Cells(15, 3).Value, "/") + 1)
      End If
    End If
  End If
  
  Set objRapidPrim = New clsRapidPrim
  If tmp(1) + tmp(2) = 0 Then
    fmt = """Å}""0.0##"
  Else
    fmt = "+0.0##;-0.0##"
  End If
  Set rpdTextData = New zwDrawCAD.TextData
  rpdTextData.Text = Format(tmp(1), fmt)
  rpdTextData.FaceName = rpObject.strItem(1)
  rpdTextData.Height = rpObject.dblHW(lngRow, 1) / rpdDoc.DrawInfo.Numerator
  rpdTextData.Width = rpObject.dblHW(lngRow, 2) / rpdDoc.DrawInfo.Numerator
  rpdTextData.BasePosition = 0
  Select Case dblAngle
    Case 0
      rpdTextData.BasePoint.x = dblBaseX + dblShift
      rpdTextData.BasePoint.y = dblBaseY + _
        (rpObject.dblHW(lngRow, 1) * 1.2) / rpdDoc.DrawInfo.Numerator
    Case 90
      rpdTextData.BasePoint.x = dblBaseX - _
        (rpObject.dblHW(lngRow, 1) * 1.2) / rpdDoc.DrawInfo.Numerator
      rpdTextData.BasePoint.y = dblBaseY + dblShift
    Case -90, 270
      rpdTextData.BasePoint.x = dblBaseX + _
        (rpObject.dblHW(lngRow, 1) * 1.2) / rpdDoc.DrawInfo.Numerator
      rpdTextData.BasePoint.y = dblBaseY - dblShift
    Case Else
      rpdTextData.BasePoint.x = dblBaseX
      rpdTextData.BasePoint.y = dblBaseY
  End Select
  rpdTextData.Angle = dblAngle
  objRapidPrim.SetUp prmObject:=rpdTextData, lngLayerNo:=CLng(Alpha(13)), _
                     lngPenStyle:=zwStyleSolid, lngPenWidth:=zwWidthThin
  rpdDoc.DataBase.Add Prim:=rpdTextData
  Set rpdTextData = Nothing
  
  If tmp(1) + tmp(2) <> 0 Then
    Set rpdTextData = New zwDrawCAD.TextData
    rpdTextData.Text = Format(tmp(2), fmt)
    rpdTextData.FaceName = rpObject.strItem(1)
    rpdTextData.Height = rpObject.dblHW(lngRow, 1) / rpdDoc.DrawInfo.Numerator
    rpdTextData.Width = rpObject.dblHW(lngRow, 2) / rpdDoc.DrawInfo.Numerator
    rpdTextData.BasePosition = 0
    Select Case dblAngle
      Case 0
        rpdTextData.BasePoint.x = dblBaseX + dblShift
        rpdTextData.BasePoint.y = dblBaseY
      Case 90
        rpdTextData.BasePoint.x = dblBaseX
        rpdTextData.BasePoint.y = dblBaseY + dblShift
      Case -90, 270
        rpdTextData.BasePoint.x = dblBaseX
        rpdTextData.BasePoint.y = dblBaseY - dblShift
      Case Else
        rpdTextData.BasePoint.x = dblBaseX
        rpdTextData.BasePoint.y = dblBaseY
    End Select
    rpdTextData.Angle = dblAngle
    objRapidPrim.SetUp prmObject:=rpdTextData, lngLayerNo:=CLng(Alpha(13)), _
                       lngPenStyle:=zwStyleSolid, lngPenWidth:=zwWidthThin
    rpdDoc.DataBase.Add Prim:=rpdTextData
    Set rpdTextData = Nothing
  End If
  Set objRapidPrim = Nothing
End With

End Sub


Public Sub Check(ByRef rpObject As clsRapidParam, ByVal lngRow As Long)

Dim i As Long, cc As Long, maxCol As Long

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
    If IsEmpty(.Cells(rpObject.lngVal(lngRow, 1), cc).Value) Then
      strOut = "Empty"
    ElseIf IsNumeric(.Cells(rpObject.lngVal(lngRow, 1), cc).Value) Then
      strOut = "Num"
    Else
      strOut = .Cells(rpObject.lngVal(lngRow, 1), cc).Value
    End If
  Else
    strOut = "Empty"
  End If
End With

End Sub


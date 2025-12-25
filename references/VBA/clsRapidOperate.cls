VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsRapidOperate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False


Option Explicit
Option Base 1
Option Compare Binary


Public Sub Run(ByVal strMode As String)

Dim objRapidDoc As clsRapidDoc
Dim objRapidLine As clsRapidLine
Dim objRapidText As clsRapidText
Dim objRapidHeader As clsRapidHeader
Dim objRapidFooter As clsRapidFooter
Dim objRapidLayout As clsRapidLayout
Dim objRapidTable As clsRapidTable
Dim objRapidDime As clsRapidDime
Dim objRapidPick As clsRapidPick
Dim objRapidGeo As clsRapidGeo
Dim objRapidEdit As clsRapidEdit
Dim objAPData As clsAssyPlotData
Dim objGPData As clsGlassPlotData
Dim objPPData As clsPressPlotData
Dim objBPData As clsBalsamPlotData

Dim i As Long, j As Long, sc As Long
Dim scAssy(4, 2) As Long
Dim scAsp(4, 2) As Long, scSph(4, 2) As Long

scAssy(1, 1) = 10: scAssy(1, 2) = 12
scAssy(2, 1) = 5: scAssy(2, 2) = 24
scAssy(3, 1) = 2: scAssy(3, 2) = 60
scAssy(4, 1) = 1: scAssy(4, 2) = 2147483647

scAsp(1, 1) = 10: scAsp(1, 2) = 12
scAsp(2, 1) = 5: scAsp(2, 2) = 24
scAsp(3, 1) = 2: scAsp(3, 2) = 60
scAsp(4, 1) = 1: scAsp(4, 2) = 2147483647

scSph(1, 1) = 10: scSph(1, 2) = 8
scSph(2, 1) = 5: scSph(2, 2) = 16
scSph(3, 1) = 2: scSph(3, 2) = 40
scSph(4, 1) = 1: scSph(4, 2) = 2147483647

Set objRapidDoc = New clsRapidDoc
Set objRapidLine = New clsRapidLine
Set objRapidText = New clsRapidText
Set objRapidDime = New clsRapidDime
Set objRapidPick = New clsRapidPick
Set objRapidGeo = New clsRapidGeo
Set objRapidEdit = New clsRapidEdit
Select Case Left(ActiveSheet.Name, 1)
  Case "G"
    If ActiveSheet.Name = "G00" Then
      If strMode <> "Glass" And _
         strMode <> "Press" Then
        Set objAPData = New clsAssyPlotData
        Set objRapidHeader = New clsRapidHeader
        Set objRapidFooter = New clsRapidFooter
        Set objRapidLayout = New clsRapidLayout
        Set objRapidTable = New clsRapidTable
        objAPData.Calc
        objProgCtrl.InitBar2 _
          lngMaxCount:=objAPData.lngMaxCount + 6
        
        If Cells(2, 6).Value = "Ž©“®" Then
          For i = 1 To UBound(scAssy, 1)
            If objAPData.dblMaxDiameter <= scAssy(i, 2) Then
              objRapidDoc.Add lngScale:=scAssy(i, 1)
              Exit For
            End If
          Next
        Else
          objRapidDoc.Add lngScale:=Val(Cells(2, 6).Value)
        End If
        With rpdDoc
          .DrawInfo.Number = .DrawInfo.Number & "_1"
          .FullName = .DrawInfo.Name & "_" & .DrawInfo.Number
        End With
        objAPData.CreateElement
        objRapidText.Add
        objRapidHeader.AddA apdObject:=objAPData
        objRapidFooter.AddA apdObject:=objAPData
        objRapidLayout.AddA apdObject:=objAPData
        
        objRapidDoc.Add lngScale:=1
        With rpdDoc
          .DrawInfo.Number = .DrawInfo.Number & "_2"
          .FullName = .DrawInfo.Name & "_" & .DrawInfo.Number
        End With
        objRapidEdit.RemoveExStd
        objRapidEdit.RemoveNoteFrame
        objRapidEdit.RemoveGlaDBFrame
        objRapidText.Add
        objRapidTable.AddA
        Set objAPData = Nothing
        Set objRapidHeader = Nothing
        Set objRapidFooter = Nothing
        Set objRapidLayout = Nothing
        Set objRapidTable = Nothing
      Else
        objProgCtrl.InitBar2 lngMaxCount:=1
        objProgCtrl.AddBar lngDivCount:=1
      End If
    Else
      If strMode <> "Assy" Then
        If InStr(ActiveSheet.Name, "ASP") = 0 And _
           strMode = "All" Then
          objProgCtrl.InitBar2 lngMaxCount:=10
        ElseIf InStr(ActiveSheet.Name, "ASP") > 0 And _
           strMode = "Press" Then
          objProgCtrl.InitBar2 lngMaxCount:=1
          objProgCtrl.AddBar lngDivCount:=1
        Else
          objProgCtrl.InitBar2 lngMaxCount:=5
        End If
        
        If Cells(2, 6).Value = "Ž©“®" Then
          If InStr(ActiveSheet.Name, "ASP") = 0 Then
            For i = 1 To UBound(scSph, 1)
              If Cells(8, 6).Value <= scSph(i, 2) Then
                sc = scSph(i, 1): Exit For
              End If
            Next
          Else
            For i = 1 To UBound(scAsp, 1)
              If Cells(8, 6).Value <= scAsp(i, 2) Then
                sc = scAsp(i, 1): Exit For
              End If
            Next
          End If
        Else
          sc = Val(Cells(2, 6).Value)
        End If
        
        If InStr(ActiveSheet.Name, "ASP") = 0 Then
          If strMode <> "All" And _
             CBool(fCheck(2)) Then
            fCheck(2) = "False"
            objRapidDoc.Add lngScale:=sc
            fCheck(2) = "True"
          Else
            objRapidDoc.Add lngScale:=sc
          End If
          
          With rpdDoc
            If strMode = "All" And _
               CBool(fCheck(2)) Then
              .DrawInfo.Number = .DrawInfo.Number & _
                                 ",Z" & Mid(ActiveSheet.Name, 2, 2)
            ElseIf strMode <> "Glass" Then
              .DrawInfo.Number = "Z" & Mid(ActiveSheet.Name, 2, 2)
            End If
            .FullName = .DrawInfo.Name & "_" & .DrawInfo.Number
          End With
        ElseIf strMode <> "Press" Then
          objRapidDoc.Add lngScale:=sc
        End If
        
        If InStr(ActiveSheet.Name, "ASP") = 0 And _
           strMode <> "Glass" Then
          Set objPPData = New clsPressPlotData
          objPPData.Calc
          objRapidEdit.RemoveExStd
          objRapidEdit.CreatePressBase
          objRapidLine.AddP ppdObject:=objPPData
          objRapidText.AddP ppdObject:=objPPData
          objRapidDime.AddP ppdObject:=objPPData
          objRapidPick.AddP ppdObject:=objPPData
          objRapidEdit.EditGlaDBBlank
          Set objPPData = Nothing
        End If
        
        If InStr(ActiveSheet.Name, "ASP") = 0 And _
           strMode = "All" And _
           Not CBool(fCheck(2)) Then
          objRapidDoc.Add lngScale:=sc
        End If
        
        If strMode <> "Press" Then
          Set objGPData = New clsGlassPlotData
          objGPData.Calc
          objRapidLine.AddG gpdObject:=objGPData
          If InStr(ActiveSheet.Name, "ASP") > 0 And _
             CBool(fCheck(4)) Then
            objRapidText.AddE gpdObject:=objGPData
          Else
            objRapidText.Add
          End If
          objRapidEdit.EditAspFrame gpdObject:=objGPData
          objRapidDime.AddG gpdObject:=objGPData
          objRapidPick.AddG gpdObject:=objGPData
          objRapidGeo.AddG gpdObject:=objGPData
          objRapidEdit.EditGlaDBBlank
          Set objGPData = Nothing
        End If
      Else
        objProgCtrl.InitBar2 lngMaxCount:=1
        objProgCtrl.AddBar lngDivCount:=1
      End If
    End If
    
  Case "B", "M"
    If strMode <> "Glass" And _
       strMode <> "Press" Then
      objProgCtrl.InitBar2 lngMaxCount:=7
      
      Set objBPData = New clsBalsamPlotData
      objBPData.PreCalc
      If Cells(2, 6).Value = "Ž©“®" Then
        For i = 1 To UBound(scSph, 1)
          If objBPData.dblMaxDiameter <= scSph(i, 2) Then
            objRapidDoc.Add lngScale:=scSph(i, 1)
            Exit For
          End If
        Next
      Else
        objRapidDoc.Add lngScale:=Val(Cells(2, 6).Value)
      End If
      objRapidEdit.RemoveExStd
      objRapidEdit.RemoveGlaDBFrame
      objBPData.CreateElement
      objRapidText.Add
      objRapidDime.AddB bpdObject:=objBPData
      objRapidPick.AddB bpdObject:=objBPData
      Set objBPData = Nothing
    Else
      objProgCtrl.InitBar2 lngMaxCount:=1
      objProgCtrl.AddBar lngDivCount:=1
    End If
    
  Case "H"
    If strMode <> "Assy" And _
       strMode <> "Press" Then
      objProgCtrl.InitBar2 lngMaxCount:=5
      
      If Cells(2, 6).Value = "Ž©“®" Then
        For i = 1 To UBound(scAsp, 1)
          If Cells(8, 6).Value <= scAsp(i, 2) Then
            objRapidDoc.Add lngScale:=scAsp(i, 1)
            Exit For
          End If
        Next
      Else
        objRapidDoc.Add lngScale:=Val(Cells(2, 6).Value)
      End If
      Set objGPData = New clsGlassPlotData
      objGPData.Calc
      objRapidEdit.RemoveGmNote
      objRapidEdit.CreateHybBase gpdObject:=objGPData
      objRapidLine.AddG gpdObject:=objGPData
      If CBool(fCheck(4)) Then
        objRapidText.AddE gpdObject:=objGPData
      Else
        objRapidText.Add
      End If
      objRapidEdit.EditAspFrame gpdObject:=objGPData
      objRapidDime.AddG gpdObject:=objGPData
      objRapidPick.AddG gpdObject:=objGPData
      Set objGPData = Nothing
    Else
      objProgCtrl.InitBar2 lngMaxCount:=1
      objProgCtrl.AddBar lngDivCount:=1
    End If
End Select
Set rpdDoc = Nothing
Set objRapidDoc = Nothing
Set objRapidLine = Nothing
Set objRapidText = Nothing
Set objRapidDime = Nothing
Set objRapidPick = Nothing
Set objRapidGeo = Nothing
Set objRapidEdit = Nothing

End Sub


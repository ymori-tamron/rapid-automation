VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsCfgFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False


Option Explicit
Option Base 1
Option Compare Binary


Public Sub OpenFile()

Dim xlWb As Excel.Workbook

Dim FileToOpen As Variant, Filter As Variant

Filter = "コンフィグファイル (*.cfg;*.txt),*.cfg;*.txt," & _
         "すべてのファイル (*.*),*.*"
FileToOpen = Application.GetOpenFilename(FileFilter:=Filter, _
                                         FilterIndex:=1, _
                                         Title:=ProjectTitle)
If FileToOpen <> False Then
  Workbooks.OpenText _
    Filename:=FileToOpen, DataType:=xlDelimited, ConsecutiveDelimiter:=False, _
    Tab:=False, Semicolon:=False, Comma:=True, Space:=False, Other:=False, _
    TextQualifier:=xlTextQualifierDoubleQuote
  Set xlWb = ActiveWorkbook
  xlWb.ActiveSheet.Copy
  xlWb.Close SaveChanges:=False
  Set xlWb = Nothing
End If

End Sub


Public Sub OverWrite()

Dim f As Integer
Dim i As Long
Dim cfgFileName As String

objError.Message intMsgNumber:=202, _
                 strAddWord1:=ActiveSheet.Name
If objError.intReturnValue = vbOK Then
  cfgFileName = ThisWorkbook.Path & "\Config\" & ActiveSheet.Name & ".cfg"
  f = FreeFile(0)
  Open cfgFileName For Output As #f
  For i = 1 To Cells(Rows.Count, 1).End(xlUp).Row
    If ActiveSheet.Name = "PressParam" Then
      Write #f, Cells(i, 1).Value, Cells(i, 2).Value, Cells(i, 3).Value, Cells(i, 4).Value, _
                Cells(i, 5).Value, Cells(i, 6).Value, Cells(i, 7).Value, Cells(i, 8).Value, _
                Cells(i, 9).Value, Cells(i, 10).Value, Cells(i, 11).Value
    Else
      Write #f, Cells(i, 1).Value, Cells(i, 2).Value, Cells(i, 3).Value, Cells(i, 4).Value, _
                Cells(i, 5).Value, Cells(i, 6).Value, Cells(i, 7).Value, Cells(i, 8).Value
    End If
  Next
  Close #f
  ActiveWorkbook.Close SaveChanges:=False
End If

End Sub


Public Sub PlotMark()

Dim rpdFrm As zwDrawCAD.Document
Dim rpdPolyData As zwDrawCAD.PolyLineData
Dim rpdDPoint As zwDrawCAD.DPoint
Dim objRapidPrim As clsRapidPrim

Dim FileToOpen As Variant, Filter As Variant
Dim f As Integer
Dim i As Long
Dim lngIndex As Long, lngRow As Long, lngCol As Long
Dim dblX As Double, dblY As Double, dblH As Double, dblW As Double
Dim strItem As String
Dim frmFileName As String, fil As String

Filter = "コンフィグファイル (*.cfg;*.txt),*.cfg;*.txt," & _
         "すべてのファイル (*.*),*.*"
FileToOpen = Application.GetOpenFilename(FileFilter:=Filter, _
                                         FilterIndex:=1, _
                                         Title:=ProjectTitle)
If FileToOpen <> False Then
  If InStrRev(FileToOpen, "\") > 0 Then
    fil = Mid(FileToOpen, InStrRev(FileToOpen, "\") + 1)
    If Left(fil, 3) = "Txt" Then
      frmFileName = ThisWorkbook.Path & "\Template\"
      If InStr(fil, "Assy") = 0 Then
        If InStr(fil, "A3E") > 0 Then
          frmFileName = frmFileName & "A3非球面E.zrd"
        ElseIf InStr(fil, "A3") > 0 Then
          frmFileName = frmFileName & "A3非球面.zrd"
        Else
          frmFileName = frmFileName & "A4球面.zrd"
        End If
      Else
        If InStr(fil, "A3") > 0 Then
          frmFileName = frmFileName & "A3配置図.zrd"
        Else
          frmFileName = frmFileName & "A4球面.zrd"
        End If
      End If
      
      Set rpdApp = CreateObject("zwDrawCAD.Application")
      With rpdApp
        .WindowState = zwWindowStateNormal
        .WindowState = zwWindowStateMaximize
        .Visible = True
        
        Set rpdFrm = .Documents.Open(sFileName:=frmFileName)
        Set rpdDoc = .Documents.Add
      End With
      
      With rpdDoc
        With .DrawInfo
          .Name = "テキスト座標(" & fil & ")"
          .Number = ""
          .SetScale lNume:=1, lDeno:=1
          If rpdFrm.DrawInfo.Sheet = zwSheetA3 Then
            .Sheet = zwSheetA3
            .Direction = zwSheetHorizontal
          Else
            .Sheet = zwSheetA4
            .Direction = zwSheetVertical
          End If
        End With
        
        .CurrentDraw.LayerNo = CLng(Alpha(10))
        .FullName = .DrawInfo.Name & ".zfd"
        .Views(0).AllDraw
        
        With .DrawLayers
          For i = 10 To 13
            If .GetLayer(lNumber:=CLng(Alpha(i))) = -1 Then
              .Add lNumber:=CLng(Alpha(i))
            End If
            
            With .Item(.GetLayer(lNumber:=CLng(Alpha(i))))
              Select Case i
                Case 10
                  .PenStyle = zwStyleSolid
                  .PenWidth = zwWidthMidThin
                Case 11
                  .PenStyle = zwStyleS1DashDot
                  .PenWidth = zwWidthThin
                Case 12, 13
                  .PenStyle = zwStyleSolid
                  .PenWidth = zwWidthThin
              End Select
              .Visible = True
            End With
          Next
          
          If .GetLayer(lNumber:=100) = -1 Then
            .Add lNumber:=100
          End If
          
          With .Item(.GetLayer(lNumber:=100))
            .PenStyle = zwStyleSolid
            .PenWidth = zwWidthMiddle
            .Visible = True
            .Comment = "レイヤ100"
          End With
        End With
      End With
      
      Set rpdDPoint = New zwDrawCAD.DPoint
      rpdDPoint.x = 0
      rpdDPoint.y = 0
      rpdFrm.DataBase.SelectAll bSelect:=True
      rpdDoc.DocCopyS doc:=rpdFrm, BasePoint:=rpdDPoint, PutPoint:=rpdDPoint, _
                      dRatio:=rpdFrm.DrawInfo.Numerator, dAngle:=0
      rpdFrm.DataBase.SelectAll bSelect:=False
      rpdFrm.Close
      rpdDoc.Activate
      
      With rpdDoc.DrawLayers
        If .GetLayer(lNumber:=999) = -1 Then
          .Add lNumber:=999
        End If
        
        With .Item(.GetLayer(lNumber:=999))
          .Color = vbRed
          .PenStyle = zwStyleSolid
          .PenWidth = zwWidthMidThin
          .Visible = True
          .Comment = "レイヤ999"
        End With
      End With
      
      Set objRapidPrim = New clsRapidPrim
      
      f = FreeFile(0)
      Open FileToOpen For Input As #f
      Do While Not EOF(f)
        Input #f, lngIndex, strItem, dblX, dblY, dblH, dblW, lngRow, lngCol
        
        rpdDPoint.x = dblX / rpdDoc.DrawInfo.Numerator
        rpdDPoint.y = dblY / rpdDoc.DrawInfo.Numerator
        
        Set rpdPolyData = New zwDrawCAD.PolyLineData
        rpdPolyData.CreateMark Type:=zwMarkCross, _
                               PutPoint:=rpdDPoint, dSize:=4
        objRapidPrim.SetUp prmObject:=rpdPolyData, lngLayerNo:=999, _
                           lngPenStyle:=zwStyleSolid, lngPenWidth:=zwWidthMidThin
        rpdDoc.DataBase.Add Prim:=rpdPolyData
        Set rpdPolyData = Nothing
      Loop
      Close #f
      
      Set rpdFrm = Nothing
      Set rpdDoc = Nothing
      Set rpdDPoint = Nothing
      Set objRapidPrim = Nothing
      
      Set rpdApp = Nothing
    Else
      objError.Message intMsgNumber:=11
    End If
  End If
End If

End Sub


VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsRapidDoc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False


Option Explicit
Option Base 1
Option Compare Binary


Public Sub Add(ByVal lngScale As Long)

Dim rpdFrm As zwDrawCAD.Document
Dim rpdDPoint As zwDrawCAD.DPoint

Dim i As Long, c(4) As Long
Dim frmFileName As String

frmFileName = ThisWorkbook.Path & "\Template\"
If ActiveSheet.Name <> "G00" Then
  If InStr(ActiveSheet.Name, "ASP") > 0 Then
    If CBool(fCheck(4)) Then
      frmFileName = frmFileName & "A3非球面E.zrd"
    Else
      frmFileName = frmFileName & "A3非球面.zrd"
    End If
  Else
    frmFileName = frmFileName & "A4球面.zrd"
  End If
Else
  If rpdDoc Is Nothing Then
    frmFileName = frmFileName & "A3配置図.zrd"
  Else
    frmFileName = frmFileName & "A4球面.zrd"
  End If
End If

With rpdApp
  Set rpdFrm = .Documents.Open(sFileName:=frmFileName)
  Set rpdDoc = .Documents.Add
End With

With rpdDoc
  With .DrawInfo
    .Name = Cells(2, 3).Value
    .Number = ActiveSheet.Name
    .SetScale lNume:=lngScale, lDeno:=1
    If rpdFrm.DrawInfo.Sheet = zwSheetA3 Then
      .Sheet = zwSheetA3
      .Direction = zwSheetHorizontal
    Else
      .Sheet = zwSheetA4
      .Direction = zwSheetVertical
    End If
  End With
  
  .CurrentDraw.LayerNo = CLng(Alpha(10))
  .FullName = .DrawInfo.Name & "_" & .DrawInfo.Number
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
        c(i - 9) = .Color
      End With
      
      lngPre(i - 9) = CLng(Alpha(i))
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
    
    If ActiveSheet.Name <> "G00" Then
      If Left(ActiveSheet.Name, 1) = "G" And _
         InStr(ActiveSheet.Name, "ASP") = 0 And _
         CBool(fCheck(2)) Then
        For i = 10 To 13
          lngPre(i - 9) = CLng(Alpha(i)) + 100
          
          If .GetLayer(lNumber:=lngPre(i - 9)) = -1 Then
            .Add lNumber:=lngPre(i - 9)
          End If
          
          With .Item(.GetLayer(lNumber:=lngPre(i - 9)))
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
            .Visible = False
            .Color = c(i - 9)
            .Comment = "レイヤ" & CStr(lngPre(i - 9))
          End With
        Next
      End If
    End If
  End With
End With

Set rpdDPoint = New zwDrawCAD.DPoint
rpdDPoint.x = 0
rpdDPoint.y = 0
rpdFrm.DataBase.SelectAll bSelect:=True
rpdDoc.DocCopyS doc:=rpdFrm, BasePoint:=rpdDPoint, PutPoint:=rpdDPoint, _
                dRatio:=rpdFrm.DrawInfo.Numerator / lngScale, dAngle:=0
rpdFrm.DataBase.SelectAll bSelect:=False
rpdFrm.Close
rpdDoc.Activate

If ActiveSheet.Name <> "G00" Then
  If Left(ActiveSheet.Name, 1) = "G" And _
     InStr(ActiveSheet.Name, "ASP") = 0 And _
     CBool(fCheck(2)) Then
    rpdDoc.DataBase.CopyLayerElement _
      lSrcLayerNo:=CLng(Alpha(12)), lDstLayerNo:=lngPre(3)
  End If
End If

Set rpdFrm = Nothing
Set rpdDPoint = Nothing

End Sub


VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsRibbonCtrl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False


Option Explicit
Option Base 1
Option Compare Binary


Public Sub SetEnabled(ByVal lngIndex As Long, ByVal blnValue As Boolean)

irEnabled(lngIndex) = blnValue
Select Case lngIndex
  Case 1
    xlIR.InvalidateControl ControlID:="Import"
  Case 2
    xlIR.InvalidateControl ControlID:="Convert"
    xlIR.InvalidateControl ControlID:="Reset"
  Case 3
    xlIR.InvalidateControl ControlID:="SwitchTolWindow"
  Case 4
    xlIR.InvalidateControl ControlID:="ImportCode"
    xlIR.InvalidateControl ControlID:="MakeSagTable"
    xlIR.InvalidateControl ControlID:="Draw"
    xlIR.InvalidateControl ControlID:="DrawAssy"
    xlIR.InvalidateControl ControlID:="DrawGlass"
    xlIR.InvalidateControl ControlID:="DrawPress"
    xlIR.InvalidateControl ControlID:="Rename"
    xlIR.InvalidateControl ControlID:="CheckRename"
    xlIR.InvalidateControl ControlID:="SwitchDataSummary"
    xlIR.InvalidateControl ControlID:="ImportRapid"
    xlIR.InvalidateControl ControlID:="OpenOLE"
  Case 5
    xlIR.InvalidateControl ControlID:="ImportGlaDB"
    xlIR.InvalidateControl ControlID:="ResetGlaDB"
  Case 6
    xlIR.InvalidateControl ControlID:="CloseGlaDB"
End Select

End Sub


Public Sub SetVisible(ByVal lngIndex As Long, ByVal blnValue As Boolean)

irVisible(lngIndex) = blnValue
Select Case lngIndex
  Case 1
    xlIR.InvalidateControl ControlID:="customGroup2"
  Case 2
    xlIR.InvalidateControl ControlID:="customGroup3"
  Case 3
    xlIR.InvalidateControl ControlID:="customGroup4"
  Case 4
    xlIR.InvalidateControl ControlID:="customGroup5"
  Case 5
    xlIR.InvalidateControl ControlID:="customGroup6"
End Select

End Sub


Public Sub SetText(ByVal strMode As String, ByVal strValue As String)

Select Case strMode
  Case "Text"
    irText = strValue
    xlIR.InvalidateControl ControlID:="CommandLine"
End Select

End Sub


Public Sub SetContent(ByVal strMode As String)

Dim k As Variant
Dim chk As Boolean
Dim i As Long
Dim xmlHeader As String, xmlItem As String
Dim xmlID As String, xmlLabel As String
Dim xmlImageMSO(3) As String

Select Case strMode
  Case "SetUp"
    For i = 1 To UBound(irContent)
      xmlItem = ""
      For Each k In Worksheets
        If k.Name <> DefaultSheet1 And _
           k.Name <> DefaultSheet2 And _
           k.Name <> DefaultSheet3 Then
          If k.Index < 10 Then
            xmlID = "Content" & CStr(i) & "_S0" & CStr(k.Index)
          Else
            xmlID = "Content" & CStr(i) & "_S" & CStr(k.Index)
          End If
          
          chk = True
          xmlLabel = k.Name
          Select Case i
            Case 1
              If k.Name = "G00" Then
                If InStr(k.Cells(3, 3).Value, ",") > 0 Then
                  xmlLabel = xmlLabel & "," & Trim(Mid(k.Cells(3, 3).Value, _
                                                 InStr(k.Cells(3, 3).Value, ",") + 1))
                End If
              ElseIf Left(k.Name, 1) <> "B" And _
                     Left(k.Name, 1) <> "M" And _
                     InStr(k.Name, "ASP") = 0 Then
                xmlLabel = xmlLabel & ",Z" & Mid(k.Name, 2, 2)
              End If
            Case 2
              If k.Name = "G00" Then
                If InStr(k.Cells(3, 3).Value, ",") > 0 Then
                  xmlLabel = xmlLabel & "," & Trim(Mid(k.Cells(3, 3).Value, _
                                                 InStr(k.Cells(3, 3).Value, ",") + 1))
                End If
              ElseIf Left(k.Name, 1) <> "B" And _
                     Left(k.Name, 1) <> "M" Then
                chk = False
              End If
            Case 3
              If k.Name = "G00" Or _
                 Left(k.Name, 1) = "B" Or _
                 Left(k.Name, 1) = "M" Then
                chk = False
              End If
            Case 4
              If k.Name = "G00" Or _
                 Left(k.Name, 1) = "B" Or _
                 Left(k.Name, 1) = "M" Or _
                 InStr(k.Name, "ASP") > 0 Then
                chk = False
              Else
                xmlLabel = "Z" & Mid(k.Name, 2, 2)
              End If
          End Select
          
          If chk Then
            xmlImageMSO(1) = "QuickTablesInsertGallery"
            If k Is ActiveSheet Then xmlImageMSO(1) = "SendCopyFlag"
            
            xmlItem = xmlItem & "<button id=" & Chr(34) & xmlID & Chr(34) & _
                                       " label=" & Chr(34) & xmlLabel & Chr(34) & _
                                       " imageMso=" & Chr(34) & xmlImageMSO(1) & Chr(34) & _
                                       " onAction=" & Chr(34) & "basCLL.Ribbon_SelectMenu" & Chr(34) & "/>"
          End If
        End If
      Next
      
      If xmlItem = "" Then
        xmlHeader = "<button id=" & Chr(34) & "Content" & CStr(i) & "_NA_" & Chr(34) & _
                           " label=" & Chr(34) & "(該当なし)" & Chr(34) & _
                           " imageMso=" & Chr(34) & "QuickTablesInsertGallery" & Chr(34) & _
                           " enabled=" & Chr(34) & "false" & Chr(34) & "/>"
      Else
        If i = 5 Then
          xmlImageMSO(2) = "QuickTablesInsertGallery"
          If ActiveSheet.Name = DefaultSheet1 Then xmlImageMSO(2) = "SendCopyFlag"
          xmlImageMSO(3) = "QuickTablesInsertGallery"
          If ActiveSheet.Name = DefaultSheet3 Then xmlImageMSO(3) = "SendCopyFlag"
          
          xmlHeader = "<button id=" & Chr(34) & "Content" & CStr(i) & "_Opt" & Chr(34) & _
                             " label=" & Chr(34) & DefaultSheet1 & Chr(34) & _
                             " imageMso=" & Chr(34) & xmlImageMSO(2) & Chr(34) & _
                             " onAction=" & Chr(34) & "basCLL.Ribbon_SelectMenu" & Chr(34) & "/>" & _
                      "<button id=" & Chr(34) & "Content" & CStr(i) & "_Tol" & Chr(34) & _
                             " label=" & Chr(34) & DefaultSheet3 & Chr(34) & _
                             " imageMso=" & Chr(34) & xmlImageMSO(3) & Chr(34) & _
                             " onAction=" & Chr(34) & "basCLL.Ribbon_SelectMenu" & Chr(34) & "/>"
        Else
          xmlHeader = "<button id=" & Chr(34) & "Content" & CStr(i) & "_Opt" & Chr(34) & _
                             " label=" & Chr(34) & "全て" & Chr(34) & _
                             " imageMso=" & Chr(34) & "QuerySelectQueryType" & Chr(34) & _
                             " onAction=" & Chr(34) & "basCLL.Ribbon_SelectMenu" & Chr(34) & "/>"
        End If
        xmlHeader = xmlHeader & "<menuSeparator id=" & Chr(34) & "Content" & CStr(i) & "_MS2" & Chr(34) & "/>"
      End If
      
      irContent(i) = "<menu xmlns=" & Chr(34) & "http://schemas.microsoft.com/office/2006/01/customui" & Chr(34) & _
                          " itemSize=" & Chr(34) & "normal" & Chr(34) & ">" & _
                     "<menuSeparator id=" & Chr(34) & "Content" & CStr(i) & "_MS1" & Chr(34)
      If i = 5 Then
        irContent(i) = irContent(i) & " title=" & Chr(34) & "切替メニュー" & Chr(34) & "/>"
      Else
        irContent(i) = irContent(i) & " title=" & Chr(34) & "選択メニュー" & Chr(34) & "/>"
      End If
      irContent(i) = irContent(i) & xmlHeader & xmlItem & "</menu>"
    Next
    
  Case "Reset"
    For i = 1 To UBound(irContent)
      irContent(i) = "<menu xmlns=" & Chr(34) & "http://schemas.microsoft.com/office/2006/01/customui" & Chr(34) & _
                          " itemSize=" & Chr(34) & "normal" & Chr(34) & ">" & _
                     "</menu>"
    Next
End Select

End Sub


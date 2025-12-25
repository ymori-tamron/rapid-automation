VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsAppEvent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False


Option Explicit
Option Base 1
Option Compare Binary

Private WithEvents xlApp As Excel.Application
Attribute xlApp.VB_VarHelpID = -1


Private Sub Class_Initialize()

Set xlApp = Excel.Application

End Sub


Private Sub Class_Terminate()

Set xlApp = Nothing

End Sub


Private Sub xlApp_WorkbookBeforeClose(ByVal Wb As Workbook, Cancel As Boolean)

If Not objGlaDB Is Nothing Then
  If Wb Is objGlaDB.xlWb_DataSummary Then
    objGlaDB.ClearSumInfo
  End If
  If Wb Is objGlaDB.xlWb_GlaDB Then
    Set objGlaDB = Nothing
  End If
End If

If Wb.Name = objWindowCtrl.strWbName Then
  Cancel = True
  objError.Message intMsgNumber:=99
End If

End Sub


Private Sub xlApp_WorkbookBeforeSave(ByVal Wb As Workbook, ByVal SaveAsUI As Boolean, Cancel As Boolean)

If Wb.Name = objWindowCtrl.strWbName Then
  Cancel = True
  objError.Message intMsgNumber:=99
End If

End Sub


Private Sub xlApp_WindowActivate(ByVal Wb As Workbook, ByVal Wn As Window)

ActivateAction

End Sub


Private Sub xlApp_SheetActivate(ByVal Sh As Object)

ActivateAction

End Sub


Private Sub xlApp_WindowDeactivate(ByVal Wb As Workbook, ByVal Wn As Window)

DeactivateAction

End Sub


Private Sub xlApp_SheetDeactivate(ByVal Sh As Object)

DeactivateAction

End Sub


Public Sub ActivateAction()

Dim k As Variant
Dim chkWn As Boolean
Dim chkWs As Integer
Dim wnName As String

chkWs = -2
If Worksheets.Count > 2 Then
  For Each k In Worksheets
    If k.Name = DefaultSheet1 Or _
       k.Name = DefaultSheet3 Then
      chkWs = chkWs + 1
    End If
  Next
End If

If chkWs = 0 Then
  With ActiveWorkbook
    wnName = .Name & ":" & TolWindow
    If ActiveWindow.Caption <> wnName Then
      If objWindowCtrl.blnPressed Then
        If .Name = objWindowCtrl.strWbName Then
          objRibbonCtrl.SetEnabled lngIndex:=3, blnValue:=True
          
          chkWn = False
          For Each k In .Windows
            If k.Caption = wnName Then chkWn = True
          Next
          If chkWn Then
            objWindowCtrl.Search
            If objWindowCtrl.lngColumn <> 0 Then
              With .Windows(wnName)
                .ScrollRow = 7
                .ScrollColumn = 3
                .SmallScroll ToRight:=objWindowCtrl.lngColumn - 3
              End With
            End If
          End If
        End If
      Else
        objRibbonCtrl.SetEnabled lngIndex:=3, blnValue:=True
      End If
      
      objRibbonCtrl.SetContent strMode:="SetUp"
      objRibbonCtrl.SetEnabled lngIndex:=4, blnValue:=True
    Else
      objRibbonCtrl.SetEnabled lngIndex:=3, blnValue:=True
      Worksheets(DefaultSheet3).Activate
    End If
  End With
End If

End Sub


Public Sub DeactivateAction()

If irEnabled(3) Then
  objRibbonCtrl.SetEnabled lngIndex:=3, blnValue:=False
End If

If irEnabled(4) Then
  objRibbonCtrl.SetContent strMode:="Reset"
  objRibbonCtrl.SetEnabled lngIndex:=4, blnValue:=False
End If

End Sub


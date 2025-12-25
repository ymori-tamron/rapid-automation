Attribute VB_Name = "basCOM"


Option Explicit
Option Base 1
Option Compare Binary


Public Sub OpenCmdWindow()

objRibbonCtrl.SetVisible lngIndex:=1, blnValue:=True
objRibbonCtrl.SetText strMode:="Text", strValue:=""

End Sub


Public Sub ClickCmdButton(ByVal strMode As String)

objRibbonCtrl.SetVisible lngIndex:=1, blnValue:=False

If strMode = "Cancel" Then irText = ""
Select Case irText
  Case "外部プログラム操作ボタンを表示"
    objRibbonCtrl.SetVisible lngIndex:=2, blnValue:=True
  Case "外部プログラム操作ボタンを非表示"
    objRibbonCtrl.SetVisible lngIndex:=2, blnValue:=False
  Case "EXCEL操作ボタンを表示"
    objRibbonCtrl.SetVisible lngIndex:=3, blnValue:=True
  Case "EXCEL操作ボタンを非表示"
    objRibbonCtrl.SetVisible lngIndex:=3, blnValue:=False
  Case "図脳RAPID操作ボタンを表示"
    objRibbonCtrl.SetVisible lngIndex:=4, blnValue:=True
  Case "図脳RAPID操作ボタンを非表示"
    objRibbonCtrl.SetVisible lngIndex:=4, blnValue:=False
  Case "環境整備ボタンを表示"
    objRibbonCtrl.SetVisible lngIndex:=5, blnValue:=True
  Case "環境整備ボタンを非表示"
    objRibbonCtrl.SetVisible lngIndex:=5, blnValue:=False
  Case "全てのボタンを表示"
    objRibbonCtrl.SetVisible lngIndex:=2, blnValue:=True
    objRibbonCtrl.SetVisible lngIndex:=3, blnValue:=True
    objRibbonCtrl.SetVisible lngIndex:=4, blnValue:=True
    objRibbonCtrl.SetVisible lngIndex:=5, blnValue:=True
  Case "全てのボタンを非表示"
    objRibbonCtrl.SetVisible lngIndex:=2, blnValue:=False
    objRibbonCtrl.SetVisible lngIndex:=3, blnValue:=False
    objRibbonCtrl.SetVisible lngIndex:=4, blnValue:=False
    objRibbonCtrl.SetVisible lngIndex:=5, blnValue:=False
  Case "Import": Import
  Case "Convert": Convert
  Case "Reset": Reset
  Case "ImportCode": ImportCode
  Case "MakeSagTable": MakeSagTable
  Case "Draw": Draw strMode:="All"
  Case "DrawAssy": Draw strMode:="Assy"
  Case "DrawGlass": Draw strMode:="Glass"
  Case "DrawPress": Draw strMode:="Press"
  Case "OpenDataSummary": OpenDataSummary
  Case "Rename": Rename
  Case "CheckRename": CheckRename
  Case "UndoRename": UndoRename
  Case "OpenTolWindow": SwitchTolWindow strMode:="Open"
  Case "CloseTolWindow": SwitchTolWindow strMode:="Close"
  Case "SetIni": SetIni
  Case "OpenManual": OpenManual
  Case "LoadUtility": LoadUtility
  Case "OpenGlaDB": OpenGlaDB
  Case "ImportGlaDB": ImportGlaDB
  Case "ResetGlaDB": ResetGlaDB
  Case "CloseGlaDB": CloseGlaDB
  Case "ImportRapid": ImportRapid
  Case "OpenOLE": OpenOLE
  Case "SaveRapid": SaveRapid
  Case "PrintRapid": PrintRapid
  Case "MakeXltxDataSummary": MakeXltx strMode:="DataSummary"
  Case "MakeXltxSagTable": MakeXltx strMode:="SagTable"
  Case "MakeXltxRenameList": MakeXltx strMode:="RenameList"
  Case "ConvertXltx": ConvertXltx
  Case "OpenCfg": OpenCfg
  Case "SaveCfg": SaveCfg
  Case "PlotCfg": PlotCfg
  Case "CloseDDTOOL": CloseDDTOOL
  Case ""
  Case Else: objError.Message intMsgNumber:=0
End Select

End Sub


Public Sub OpenGlaDB()

If objGlaDB Is Nothing Then
  objError.CheckGlaDBExists
  If objError.intCode = 0 Then
    Set objGlaDB = New clsGlaDB
    objGlaDB.Load
  End If
Else
  objGlaDB.Load
End If

End Sub


Public Sub ImportGlaDB()

objError.CheckGlaDBStatus
If objError.intCode = 0 Then
  objGlaDB.PickUp
End If

End Sub


Public Sub ResetGlaDB()

If Not objGlaDB Is Nothing Then
  objGlaDB.ClearSumInfo
End If

End Sub


Public Sub CloseGlaDB()

If Not objGlaDB Is Nothing Then
  If Not objGlaDB.xlWb_GlaDB Is Nothing Then
    objGlaDB.xlWb_GlaDB.Close SaveChanges:=False
  End If
End If

End Sub


Public Sub ImportRapid()

Dim objRapidCtrl As clsRapidCtrl

Dim k As Variant

Set rpdApp = CreateObject("zwDrawCAD.Application")
objError.CheckRapidMatch
If objError.intCode = 0 Then
  Set objRapidCtrl = New clsRapidCtrl
  objRapidCtrl.TemporaryFolder strMode:="SetUp"
  
  Set objProgCtrl = New clsProgCtrl
  frmProg.Caption = ProjectTitle & " 只今読込中"
  If ActiveSheet.Name <> DefaultSheet1 And _
     ActiveSheet.Name <> DefaultSheet3 Then
    objProgCtrl.InitBar1 lngMaxCount:=1
    objRapidCtrl.RapidToOLE
  Else
    objProgCtrl.InitBar1 lngMaxCount:=Worksheets.Count - 2
    For Each k In Worksheets
      If k.Name <> DefaultSheet1 And _
         k.Name <> DefaultSheet3 Then
        k.Activate: DoEvents
        objRapidCtrl.RapidToOLE
      End If
    Next
  End If
  Set objProgCtrl = Nothing
  
  With frmMsg
    .Show
    .Image1.Visible = True
    .Label1.Caption = "一時ファイルを削除しています....."
    .Repaint
    
    objRapidCtrl.TemporaryFolder strMode:="CleanUp"
    DoEvents
    
    .Image1.Visible = False
    .Image2.Visible = True
    .Label1.Caption = "図面データの読込が完了しました"
    .Repaint
  End With
  apiMsgBeep uType:=vbInformation
  apiSleep dwMSec:=1500
  Unload Object:=frmMsg
  
  Set objRapidCtrl = Nothing
End If
Set rpdApp = Nothing

End Sub


Public Sub OpenOLE()

Dim objRapidCtrl As clsRapidCtrl

Dim k As Variant

objError.CheckOLE
If objError.intCode = 0 Then
  Set rpdApp = CreateObject("zwDrawCAD.Application")
  With rpdApp
    .WindowState = zwWindowStateNormal
    .WindowState = zwWindowStateMaximize
    .Visible = True
  End With
  
  Set objRapidCtrl = New clsRapidCtrl
  If ActiveSheet.Name <> DefaultSheet1 And _
     ActiveSheet.Name <> DefaultSheet3 Then
    objRapidCtrl.OLEToRapid
  Else
    For Each k In Worksheets
      If k.Name <> DefaultSheet1 And _
         k.Name <> DefaultSheet3 Then
        k.Activate: DoEvents
        objRapidCtrl.OLEToRapid
      End If
    Next
  End If
  Set objRapidCtrl = Nothing
  Set rpdApp = Nothing
End If

End Sub


Public Sub SaveRapid()

Dim objRapidCtrl As clsRapidCtrl

Set rpdApp = CreateObject("zwDrawCAD.Application")
objError.CheckRapidFName
If objError.intCode = 0 Then
  Set objRapidCtrl = New clsRapidCtrl
  objRapidCtrl.AllSave
  Set objRapidCtrl = Nothing
End If
Set rpdApp = Nothing

End Sub


Public Sub PrintRapid()

Dim objRapidCtrl As clsRapidCtrl

Set rpdApp = CreateObject("zwDrawCAD.Application")
objError.CheckRapidExists
If objError.intCode = 0 Then
  Set objRapidCtrl = New clsRapidCtrl
  objRapidCtrl.AllPrint
  Set objRapidCtrl = Nothing
End If
Set rpdApp = Nothing

End Sub


Public Sub MakeXltx(ByVal strMode As String)

Dim objXltxFile As clsXltxFile

Set objXltxFile = New clsXltxFile
Select Case strMode
  Case "DataSummary": objXltxFile.DataSummary
  Case "SagTable": objXltxFile.SagTable
  Case "RenameList": objXltxFile.RenameList
End Select
Set objXltxFile = Nothing

End Sub


Public Sub ConvertXltx()

Dim objXltxFile As clsXltxFile

Set objXltxFile = New clsXltxFile
objXltxFile.OldToNew
Set objXltxFile = Nothing

End Sub


Public Sub OpenCfg()

Dim objCfgFile As clsCfgFile

Set objCfgFile = New clsCfgFile
objCfgFile.OpenFile
Set objCfgFile = Nothing

End Sub


Public Sub SaveCfg()

Dim objCfgFile As clsCfgFile

objError.CheckCFG
If objError.intCode = 0 Then
  Set objCfgFile = New clsCfgFile
  objCfgFile.OverWrite
  Set objCfgFile = Nothing
End If

End Sub


Public Sub PlotCfg()

Dim objCfgFile As clsCfgFile

Set objCfgFile = New clsCfgFile
objCfgFile.PlotMark
Set objCfgFile = Nothing

End Sub


Public Sub CloseDDTOOL()

ThisWorkbook.Close SaveChanges:=False

End Sub


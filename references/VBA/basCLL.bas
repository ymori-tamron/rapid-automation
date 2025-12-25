Attribute VB_Name = "basCLL"


Option Explicit
Option Base 1
Option Compare Binary


Public Sub Ribbon_Initialize(Ribbon As IRibbonUI)

Dim objIniFile As clsIniFile

Dim i As Long

Set xlIR = Ribbon

irEnabled(1) = True
irEnabled(2) = False
irEnabled(3) = False
irEnabled(4) = False
irEnabled(5) = False
irEnabled(6) = False

Set objIniFile = New clsIniFile
objIniFile.FileToArray
Set objIniFile = Nothing
irVisible(1) = False
irVisible(2) = CBool(vCheck(1))
irVisible(3) = CBool(vCheck(2))
irVisible(4) = CBool(vCheck(3))
irVisible(5) = CBool(vCheck(4))

irText = ""

For i = 1 To UBound(irContent)
  irContent(i) = ""
Next

End Sub


Public Sub Ribbon_ClickButton(Control As IRibbonControl)

Select Case Control.ID
  Case "Import": Import
  Case "Convert": Convert
  Case "Reset": Reset
  Case "ImportCode": ImportCode
  Case "MakeSagTable": MakeSagTable
  Case "OpenDataSummary": OpenDataSummary
  Case "Rename": Rename
  Case "CheckRename": CheckRename
  Case "UndoRename": UndoRename
  Case "SetIni": SetIni
  Case "OpenManual": OpenManual
  Case "LoadUtility": LoadUtility
  Case "OpenGlaDB": OpenGlaDB
  Case "ImportGlaDB": ImportGlaDB
  Case "ResetGlaDB": ResetGlaDB
  Case "CloseGlaDB": CloseGlaDB
  Case "OpenCmdWindow": OpenCmdWindow
  Case "OKButton": ClickCmdButton strMode:="OK"
  Case "CancelButton": ClickCmdButton strMode:="Cancel"
  Case "SaveRapid": SaveRapid
  Case "PrintRapid": PrintRapid
  Case "MakeXltxDataSummary": MakeXltx strMode:="DataSummary"
  Case "MakeXltxSagTable": MakeXltx strMode:="SagTable"
  Case "MakeXltxRenameList": MakeXltx strMode:="RenameList"
  Case "ConvertXltx": ConvertXltx
  Case "OpenCfg": OpenCfg
  Case "SaveCfg": SaveCfg
  Case "PlotCfg": PlotCfg
End Select

End Sub


Public Sub Ribbon_SelectMenu(Control As IRibbonControl)

Select Case Right(Control.ID, 3)
  Case "Opt": Worksheets(DefaultSheet1).Activate
  Case "Tol": Worksheets(DefaultSheet3).Activate
  Case "NA_":
  Case Else: Worksheets(CLng(Right(Control.ID, 2))).Activate
End Select
DoEvents

Select Case Left(Control.ID, 8)
  Case "Content1": Draw strMode:="All"
  Case "Content2": Draw strMode:="Assy"
  Case "Content3": Draw strMode:="Glass"
  Case "Content4": Draw strMode:="Press"
  Case "Content6": ImportRapid
  Case "Content7": OpenOLE
End Select

End Sub


Public Sub Ribbon_PressButton(Control As IRibbonControl, Pressed As Boolean)

Select Case Pressed
  Case True: SwitchTolWindow strMode:="Open"
  Case False: SwitchTolWindow strMode:="Close"
End Select

End Sub


Public Sub Ribbon_EnterText(Control As IRibbonControl, Text As String)

irText = Text

End Sub


Public Sub Ribbon_SetEnabled(Control As IRibbonControl, ByRef Enabled)

Select Case Control.ID
  Case "Import"
    Enabled = irEnabled(1)
  Case "Convert", "Reset"
    Enabled = irEnabled(2)
  Case "SwitchTolWindow"
    Enabled = irEnabled(3)
  Case "ImportCode", "MakeSagTable", _
       "Draw", "DrawAssy", "DrawGlass", "DrawPress", _
       "Rename", "CheckRename", "SwitchDataSummary", _
       "ImportRapid", "OpenOLE"
    Enabled = irEnabled(4)
  Case "ImportGlaDB", "ResetGlaDB"
    Enabled = irEnabled(5)
  Case "CloseGlaDB"
    Enabled = irEnabled(6)
End Select

End Sub


Public Sub Ribbon_SetVisible(Control As IRibbonControl, ByRef Visible)

Select Case Control.ID
  Case "customGroup2"
    Visible = irVisible(1)
  Case "customGroup3"
    Visible = irVisible(2)
  Case "customGroup4"
    Visible = irVisible(3)
  Case "customGroup5"
    Visible = irVisible(4)
  Case "customGroup6"
    Visible = irVisible(5)
End Select

End Sub


Public Sub Ribbon_SetText(Control As IRibbonControl, ByRef Text)

Text = irText

End Sub


Public Sub Ribbon_SetContent(Control As IRibbonControl, ByRef Content)

Select Case Control.ID
  Case "Draw"
    Content = irContent(1)
  Case "DrawAssy"
    Content = irContent(2)
  Case "DrawGlass"
    Content = irContent(3)
  Case "DrawPress"
    Content = irContent(4)
  Case "SwitchDataSummary"
    Content = irContent(5)
  Case "ImportRapid"
    Content = irContent(6)
  Case "OpenOLE"
    Content = irContent(7)
End Select

End Sub


Public Sub Ribbon_SetPressed(Control As IRibbonControl, ByRef ReturnValue)

ReturnValue = objWindowCtrl.blnPressed

End Sub


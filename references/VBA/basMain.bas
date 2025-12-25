Attribute VB_Name = "basMain"


Option Explicit
Option Base 1
Option Compare Binary

#If VBA7 Then
  Public Declare PtrSafe Sub apiSleep Lib "KERNEL32.dll" Alias "Sleep" _
                                      (ByVal dwMSec As LongPtr)
  Public Declare PtrSafe Sub apiMsgBeep Lib "USER32.dll" Alias "MessageBeep" _
                                        (ByVal uType As LongPtr)
#Else
  Public Declare Sub apiSleep Lib "KERNEL32.dll" Alias "Sleep" _
                              (ByVal dwMSec As Long)
  Public Declare Sub apiMsgBeep Lib "USER32.dll" Alias "MessageBeep" _
                                (ByVal uType As Long)
#End If

Public xlIR As IRibbonUI
Public irEnabled(6) As Boolean
Public irVisible(5) As Boolean
Public irText As String
Public irContent(7) As String

Public rpdApp As zwDrawCAD.Application
Public rpdDoc As zwDrawCAD.Document
Public objAppEvent As clsAppEvent
Public objError As clsError
Public objProgCtrl As clsProgCtrl
Public objRibbonCtrl As clsRibbonCtrl
Public objWindowCtrl As clsWindowCtrl
Public objGlaDB As clsGlaDB

Public lngSur As Long, lngAsp As Long, lngZoo As Long
Public lngGla As Long, lngBal As Long
Public lngPre(4) As Long
Public mdlNum As String, curWb As String
Public DDT(99) As String, Alpha(16) As String
Public fCheck(5) As String, vCheck(4) As String

Public Const ProjectTitle As String = "図面作成TOOL"
Public Const DefaultSheet1 As String = "光学データサマリ"
Public Const DefaultSheet2 As String = "硝子データサマリ"
Public Const DefaultSheet3 As String = "公差データサマリ"
Public Const MainWindow As String = "メインウィンドウ"
Public Const TolWindow As String = "公差ウィンドウ"
Public Const PrimaryOLEObject As String = "PrimaryDraw"
Public Const SecondaryOLEObject As String = "SecondaryDraw"


Public Sub Import()

Dim objOpticalData As clsOpticalData

Dim FileToOpen As Variant, Filter As Variant
Dim i As Long
Dim winT As Double, winL As Double
Dim winH As Double, winW As Double

lngSur = 0: lngAsp = 0: lngZoo = 0
lngGla = 0: lngBal = 0
mdlNum = "": curWb = ""
For i = 1 To UBound(DDT)
  DDT(i) = ""
Next

Filter = "シーケンスファイル (*.seq),*.seq," & _
         "すべてのファイル (*.*),*.*"
FileToOpen = Application.GetOpenFilename(FileFilter:=Filter, _
                                         FilterIndex:=1, _
                                         Title:=ProjectTitle)
If FileToOpen <> False Then
  Workbooks.Open _
    Filename:=ThisWorkbook.Path & "\Template\図面データサマリ.xltx"
  DoEvents
  
  With ActiveWorkbook
    With .BuiltinDocumentProperties("Creation Date")
      .Value = Now
      curWb = .Value
    End With
    
    If CDbl(Application.Version) >= 15 Then
      With .Windows(1)
        .WindowState = xlMaximized
        winT = .Top
        winL = .Left
        winH = .Height
        winW = .Width
        
        .WindowState = xlNormal
        .Top = winT + 6
        .Left = winL + 6
        .Height = winH - 12
        .Width = winW - 12
      End With
    Else
      .Windows.Arrange _
        ArrangeStyle:=xlArrangeStyleTiled, ActiveWorkbook:=True
    End If
  End With
  DoEvents
  
  Set objOpticalData = New clsOpticalData
  objOpticalData.Add opFileName:=CStr(FileToOpen)
  Set objOpticalData = Nothing
  
  DoEvents
  apiSleep dwMSec:=200
  
  Convert
End If

End Sub


Public Sub Convert()

Dim objIniFile As clsIniFile
Dim objGlassData As clsGlassData
Dim objToleranceData As clsToleranceData

Dim i As Long

objError.CheckConvert
If objError.intCode = 0 Then
  Set objIniFile = New clsIniFile
  objIniFile.FileToArray
  Set objIniFile = Nothing
  
  With Application
    .ScreenUpdating = False
    
    Set objGlassData = New clsGlassData
    Set objToleranceData = New clsToleranceData
    If lngGla > 0 Then
      Worksheets(DefaultSheet2).Copy _
        Before:=Worksheets(DefaultSheet3)
      ActiveSheet.Name = "G00"
      objGlassData.AddA
      
      For i = 1 To lngGla
        Worksheets(DefaultSheet2).Copy _
          Before:=Worksheets(DefaultSheet3)
        If i < 10 Then
          ActiveSheet.Name = "G0" & CStr(i)
        Else
          ActiveSheet.Name = "G" & CStr(i)
        End If
        objGlassData.AddG
        objToleranceData.AddG
        
        If objGlassData.intHyb > 0 Then
          Worksheets(DefaultSheet2).Copy _
            Before:=Worksheets(DefaultSheet3)
          If i < 10 Then
            ActiveSheet.Name = "H0" & CStr(i)
          Else
            ActiveSheet.Name = "H" & CStr(i)
          End If
          objGlassData.AddG
          objToleranceData.AddG
          
          Worksheets(DefaultSheet2).Copy _
            Before:=Worksheets(DefaultSheet3)
          If i < 10 Then
            ActiveSheet.Name = "M0" & CStr(i)
          Else
            ActiveSheet.Name = "M" & CStr(i)
          End If
          objGlassData.AddM
        End If
      Next
    End If
    
    If lngBal > 0 Then
      For i = 1 To lngBal
        Worksheets(DefaultSheet2).Copy _
          Before:=Worksheets(DefaultSheet3)
        If i < 10 Then
          ActiveSheet.Name = "B0" & CStr(i)
        Else
          ActiveSheet.Name = "B" & CStr(i)
        End If
        objGlassData.AddB
      Next
    End If
    Set objGlassData = Nothing
    Set objToleranceData = Nothing
    
    .DisplayAlerts = False
    Worksheets(DefaultSheet2).Delete
    .DisplayAlerts = True
    
    Worksheets(DefaultSheet1).Activate
    .ScreenUpdating = True
  End With
  
  Reset
  With frmMsg
    .Show
    .Image2.Visible = True
    .Label1.Caption = "光学データの読込が完了しました"
    .Repaint
  End With
  apiMsgBeep uType:=vbInformation
  apiSleep dwMSec:=1500
  Unload Object:=frmMsg
End If

End Sub


Public Sub Reset()

objRibbonCtrl.SetEnabled lngIndex:=1, blnValue:=True
objRibbonCtrl.SetEnabled lngIndex:=2, blnValue:=False

End Sub


Public Sub ImportCode()

Dim objMaterialCode As clsMaterialCode

Dim k As Variant

objError.CheckSummary
If objError.intCode = 0 Then
  objError.Message intMsgNumber:=201
  If objError.intReturnValue = vbOK Then
    Set objMaterialCode = New clsMaterialCode
    objMaterialCode.ReadCode
    If objMaterialCode.blnStatus Then
      Application.ScreenUpdating = False
      
      For Each k In Worksheets
        If k.Name <> DefaultSheet1 And _
           k.Name <> DefaultSheet3 Then
          k.Activate
          objMaterialCode.WriteCode
        End If
      Next
      
      Worksheets(DefaultSheet1).Activate
      Application.ScreenUpdating = True
      
      With frmMsg
        .Show
        .Image2.Visible = True
        .Label1.Caption = "品目コードの読込が完了しました"
        .Repaint
      End With
      apiMsgBeep uType:=vbInformation
      apiSleep dwMSec:=1500
      Unload Object:=frmMsg
    End If
    Set objMaterialCode = Nothing
  End If
End If

End Sub


Public Sub MakeSagTable()

Dim objSagTable As clsSagTable

Dim k As Variant

objError.CheckSummary
If objError.intCode = 0 Then
  Application.ScreenUpdating = False
  
  Set objSagTable = New clsSagTable
  mdlNum = Mid(Worksheets(DefaultSheet1).Cells(2, 4).Value, 5)
  For Each k In Worksheets
    If k.Name <> DefaultSheet1 And _
       k.Name <> DefaultSheet3 Then
      If InStr(k.Name, "ASP") > 0 Then
        k.Activate
        objSagTable.PickUp
      End If
    End If
  Next
  
  Worksheets(DefaultSheet1).Activate
  
  objSagTable.Export
  Set objSagTable = Nothing
  
  Application.ScreenUpdating = True
End If

End Sub


Public Sub Draw(ByVal strMode As String)

Dim objIniFile As clsIniFile
Dim objRapidOperate As clsRapidOperate

Dim k As Variant
Dim i As Long

objError.CheckSummary
If objError.intCode = 0 Then
  Set objIniFile = New clsIniFile
  objIniFile.FileToArray
  Set objIniFile = Nothing
  
  For i = 1 To UBound(lngPre)
    lngPre(i) = 0
  Next
  
  Set rpdApp = CreateObject("zwDrawCAD.Application")
  If CBool(fCheck(1)) Then
    With rpdApp
      .WindowState = zwWindowStateNormal
      .WindowState = zwWindowStateMaximize
      .Visible = True
    End With
  End If
  
  Set objProgCtrl = New clsProgCtrl
  frmProg.Caption = ProjectTitle & " 只今製図中"
  Set objRapidOperate = New clsRapidOperate
  If ActiveSheet.Name <> DefaultSheet1 And _
     ActiveSheet.Name <> DefaultSheet3 Then
    objProgCtrl.InitBar1 lngMaxCount:=1
    objRapidOperate.Run strMode:=strMode
  Else
    objProgCtrl.InitBar1 lngMaxCount:=Worksheets.Count - 2
    For Each k In Worksheets
      If k.Name <> DefaultSheet1 And _
         k.Name <> DefaultSheet3 Then
        k.Activate: DoEvents
        objRapidOperate.Run strMode:=strMode
      End If
    Next
  End If
  Set objRapidOperate = Nothing
  Set objProgCtrl = Nothing
  
  With frmMsg
    .Show
    .Image2.Visible = True
    .Label1.Caption = "製図が完了しました"
    .Repaint
  End With
  apiMsgBeep uType:=vbInformation
  apiSleep dwMSec:=1500
  Unload Object:=frmMsg
  
  If Not CBool(fCheck(1)) Then
    With rpdApp
      .WindowState = zwWindowStateNormal
      .WindowState = zwWindowStateMaximize
      .Visible = True
    End With
  End If
  Set rpdApp = Nothing
End If

End Sub


Public Sub OpenDataSummary()

Dim objFuncLinkCtrl As clsFuncLinkCtrl

Set objFuncLinkCtrl = New clsFuncLinkCtrl
objFuncLinkCtrl.OpenAndRemove
Set objFuncLinkCtrl = Nothing

End Sub


Public Sub Rename()

Dim fDialog As FileDialog
Dim objFileNameCtrl As clsFileNameCtrl

Dim tarFolderName As String

objError.CheckSummary
If objError.intCode = 0 Then
  Set fDialog = Application.FileDialog(fileDialogType:= _
                                       msoFileDialogFolderPicker)
  With fDialog
    .Title = ProjectTitle
    .ButtonName = "対象フォルダに指定"
    If .Show = -1 Then
      tarFolderName = .SelectedItems(1) & "\"
    Else
      tarFolderName = ""
    End If
  End With
  Set fDialog = Nothing
  Application.ScreenUpdating = True
  
  If tarFolderName <> "" Then
    Set objFileNameCtrl = New clsFileNameCtrl
    objFileNameCtrl.RenameBySummary tarFolderName:=tarFolderName
    Set objFileNameCtrl = Nothing
  End If
End If

End Sub


Public Sub CheckRename()

Dim objFileNameCtrl As clsFileNameCtrl

objError.CheckSummary
If objError.intCode = 0 Then
  Set objFileNameCtrl = New clsFileNameCtrl
  objFileNameCtrl.ListRename
  Set objFileNameCtrl = Nothing
End If

End Sub


Public Sub UndoRename()

Dim objFileNameCtrl As clsFileNameCtrl

Dim FileToOpen As Variant, Filter As Variant

objError.Message intMsgNumber:=102

Filter = "ログファイル (*.log),*.log," & _
         "すべてのファイル (*.*),*.*"
FileToOpen = Application.GetOpenFilename(FileFilter:=Filter, _
                                         FilterIndex:=1, _
                                         Title:=ProjectTitle)
If FileToOpen <> False Then
  Set objFileNameCtrl = New clsFileNameCtrl
  objFileNameCtrl.RenameByLogFile tarLogFileName:=CStr(FileToOpen)
  Set objFileNameCtrl = Nothing
End If

End Sub


Public Sub SwitchTolWindow(ByVal strMode As String)

objError.CheckSummary
If objError.intCode = 0 Then
  Select Case strMode
    Case "Open": objWindowCtrl.OpenWn
    Case "Close": objWindowCtrl.CloseWn
  End Select
End If

End Sub


Public Sub SetIni()

frmIni.Show

End Sub


Public Sub OpenManual()

Dim shObject As Object

objError.CheckManualExists
If objError.intCode = 0 Then
  Set shObject = CreateObject("WScript.Shell")
  shObject.Run Chr(34) & ThisWorkbook.Path & _
                         "\取扱説明書.pdf" & Chr(34), 5, False
  Set shObject = Nothing
End If

End Sub


Public Sub LoadUtility()

objError.CheckUtilityExists
If objError.intCode = 0 Then
  Workbooks.Open _
    Filename:=ThisWorkbook.Path & "\Utility\Utility.xlam"
End If

End Sub


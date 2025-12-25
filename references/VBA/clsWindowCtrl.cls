VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsWindowCtrl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False


Option Explicit
Option Base 1
Option Compare Binary

Private blnOutPre As Boolean
Private lngOutCol As Long
Private strOutWbn As String

Private curS As Long
Private curT As Double, curL As Double
Private curH As Double, curW As Double


Private Sub Class_Initialize()

blnOutPre = False
lngOutCol = 0
strOutWbn = ""

curS = xlMaximized
curT = 1: curL = 1
curH = 100: curW = 100

End Sub


Public Property Get blnPressed() As Boolean

blnPressed = blnOutPre

End Property


Public Property Get lngColumn() As Long

lngColumn = lngOutCol

End Property


Public Property Get strWbName() As String

strWbName = strOutWbn

End Property


Public Sub Search()

Dim i As Long, maxCol As Long

With Worksheets(DefaultSheet3)
  lngOutCol = 0
  maxCol = .Cells(6, .Columns.Count).End(xlToLeft).Column
  If maxCol > 2 Then
    For i = 3 To maxCol Step 3
      If .Cells(6, i).Value = ActiveSheet.Name Then
        lngOutCol = i: Exit For
      End If
    Next
  End If
End With

End Sub


Public Sub OpenWn()

Dim wnName(2) As String

With ActiveWorkbook
  strOutWbn = .Name
  wnName(1) = .Name & ":" & MainWindow
  wnName(2) = .Name & ":" & TolWindow
  
  With .Windows(1)
    curS = .WindowState
    .WindowState = xlNormal
    curT = .Top
    curL = .Left
    curH = .Height
    curW = .Width
    
    .Caption = wnName(1)
    .Width = curW - 228
  End With
  
  .NewWindow
  With .Windows(1)
    .Caption = wnName(2)
    .Top = curT
    .Left = curL + (curW - 228)
    .Height = curH
    .Width = 228
    
    Worksheets(DefaultSheet3).Activate
    .SplitRow = 6
    .SplitColumn = 2
    .FreezePanes = True
  End With
  
  .Windows(wnName(1)).Activate
End With

blnOutPre = True
objAppEvent.ActivateAction

End Sub


Public Sub CloseWn()

Dim k As Variant
Dim chkWn(2) As Boolean
Dim wnName(2) As String

blnOutPre = False

With ActiveWorkbook
  strOutWbn = ""
  wnName(1) = .Name & ":" & MainWindow
  wnName(2) = .Name & ":" & TolWindow
  
  chkWn(1) = False: chkWn(2) = False
  For Each k In .Windows
    Select Case k.Caption
      Case wnName(1): chkWn(1) = True
      Case wnName(2): chkWn(2) = True
    End Select
  Next
  
  If chkWn(1) Then
    .Windows(wnName(1)).Activate
    If chkWn(2) Then .Windows(wnName(2)).Close
  ElseIf chkWn(2) Then
    .Windows(wnName(2)).Activate
  End If
  
  With .Windows(1)
    .WindowState = xlNormal
    
    .Caption = ActiveWorkbook.Name
    .Top = curT
    .Left = curL
    .Height = curH
    .Width = curW
    
    .WindowState = curS
  End With
End With

objAppEvent.ActivateAction

End Sub


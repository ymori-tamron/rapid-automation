VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsFuncLinkCtrl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False


Option Explicit
Option Base 1
Option Compare Binary


Public Sub OpenAndRemove()

Dim xlWb As Excel.Workbook
Dim fDialog As FileDialog

Dim k As Variant, curLinks As Variant
Dim chkWb As Boolean
Dim chkWs As Integer
Dim i As Long
Dim txt As String

Set fDialog = Application.FileDialog(fileDialogType:= _
                                     msoFileDialogOpen)
With fDialog
  .Title = ProjectTitle
  .ButtonName = "開く"
  .FilterIndex = 13
  .AllowMultiSelect = False
  If .Show = -1 Then
    chkWb = True
    If Workbooks.Count > 0 Then
      For Each k In Workbooks
        If k.FullName = .SelectedItems(1) Then
          k.Activate
          chkWb = False
          Exit For
        End If
      Next
    End If
    If chkWb Then
      Workbooks.Open _
        Filename:=.SelectedItems(1), UpdateLinks:=0
      DoEvents
    End If
    
    Set xlWb = ActiveWorkbook
  Else
    Set xlWb = Nothing
  End If
End With
Set fDialog = Nothing
Application.ScreenUpdating = True

If Not xlWb Is Nothing Then
  Application.ScreenUpdating = False
  
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
    curLinks = xlWb.LinkSources(xlExcelLinks)
    If Not IsEmpty(curLinks) Then
      txt = "関数リンクは全て有効です"
      For i = 1 To UBound(curLinks)
        If curLinks(i) <> ThisWorkbook.FullName And _
           InStr(curLinks(i), ProjectTitle) > 0 Then
          xlWb.ChangeLink _
            Name:=curLinks(i), _
            NewName:=ThisWorkbook.FullName, _
            Type:=xlLinkTypeExcelLinks
          txt = "無効な関数リンクを修正しました"
        End If
      Next
    Else
      txt = "関数リンクが見当たりません"
    End If
    
    With frmMsg
      .Show
      .Image2.Visible = True
      .Label1.Caption = txt
      .Repaint
    End With
    apiMsgBeep uType:=vbInformation
    apiSleep dwMSec:=1500
    Unload Object:=frmMsg
  Else
    objError.Message intMsgNumber:=5
    If chkWb Then xlWb.Close SaveChanges:=False
  End If
  Set xlWb = Nothing
  
  Application.ScreenUpdating = True
End If

End Sub


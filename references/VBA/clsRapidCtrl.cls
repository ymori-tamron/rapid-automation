VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsRapidCtrl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False


Option Explicit
Option Base 1
Option Compare Binary

Private lngMC As Long
Private tmpFolder As String, zfdFileName(99) As String


Public Sub RapidToOLE()

Dim xlOLE As Excel.OLEObject

Dim k As Variant
Dim i As Long
Dim sc As Double
Dim lp(2) As Double, hp(2) As Double, wp(2) As Double
Dim prtNum As String
Dim savFileName(2) As String, oleName(2) As String

If ActiveSheet.Name = "G00" Then
  lp(1) = 636.75: lp(2) = 1069.5
  hp(1) = 297.75: hp(2) = 371.25
  wp(1) = 429.75: wp(2) = 252
ElseIf InStr(ActiveSheet.Name, "ASP") > 0 Then
  lp(1) = 522: lp(2) = 954.75
  hp(1) = 297.75: hp(2) = 371.25
  wp(1) = 429.75: wp(2) = 252
Else
  lp(1) = 522: lp(2) = 777
  hp(1) = 371.25: hp(2) = 371.25
  wp(1) = 252: wp(2) = 252
End If
mdlNum = Cells(2, 3).Value
prtNum = ActiveSheet.Name
savFileName(1) = ""
savFileName(2) = ""
oleName(1) = PrimaryOLEObject
oleName(2) = SecondaryOLEObject

objProgCtrl.InitBar2 lngMaxCount:=6
With rpdApp.Documents
  objProgCtrl.AddBar lngDivCount:=1
  For i = 0 To .Count - 1
    With .Item(i)
      If .DrawInfo.Name = mdlNum Then
        If savFileName(1) = "" Then
          If .DrawInfo.Number = prtNum & "_1" Or _
             prtNum <> "G00" And _
             Left(.DrawInfo.Number, 3) = Left(prtNum, 3) Then
            savFileName(1) = tmpFolder & .DrawInfo.Name & "_" & _
                                         .DrawInfo.Number & ".zfd"
            .Activate
            .SaveAs sFileName:=savFileName(1)
            lngMC = lngMC + 1
            zfdFileName(lngMC) = savFileName(1)
            Exit For
          End If
        End If
      End If
    End With
  Next
  
  objProgCtrl.AddBar lngDivCount:=1
  For i = 0 To .Count - 1
    With .Item(i)
      If .DrawInfo.Name = mdlNum Then
        If savFileName(2) = "" Then
          If Left(prtNum, 1) = "G" Then
            If .DrawInfo.Number = prtNum & "_2" Or _
               .DrawInfo.Number = "Z" & Mid(prtNum, 2, 2) Then
              savFileName(2) = tmpFolder & .DrawInfo.Name & "_" & _
                                           .DrawInfo.Number & ".zfd"
              .Activate
              .SaveAs sFileName:=savFileName(2)
              lngMC = lngMC + 1
              zfdFileName(lngMC) = savFileName(2)
              Exit For
            End If
          End If
        End If
      End If
    End With
  Next
End With

Application.ScreenUpdating = False
For i = 1 To 2
  If savFileName(i) <> "" Then
    objProgCtrl.AddBar lngDivCount:=1
    If ActiveSheet.OLEObjects.Count > 0 Then
      For Each k In ActiveSheet.OLEObjects
        With k
          If .Name = oleName(i) Then
            .Name = .Name & "_1"
            .Top = .Top + 6
            .Left = .Left + 6
          ElseIf InStr(.Name, "_") > 0 Then
            If Mid(.Name, 1, InStr(.Name, "_") - 1) = oleName(i) Then
              .Name = Mid(.Name, 1, InStr(.Name, "_")) & _
                      CStr(CLng(Mid(.Name, InStr(.Name, "_") + 1)) + 1)
              .Top = .Top + 6
              .Left = .Left + 6
            End If
          End If
        End With
      Next
    End If
    
    objProgCtrl.AddBar lngDivCount:=1
    Set xlOLE = _
      ActiveSheet.OLEObjects.Add(Filename:=savFileName(i), _
                                 Link:=False, _
                                 DisplayAsIcon:=False)
    With xlOLE
      .Name = oleName(i)
      .Top = 12.75
      .Left = lp(i)
      
      .ShapeRange.LockAspectRatio = msoFalse
      
      sc = hp(i) / .Height
      .Height = .Height * sc
      .Width = .Width * sc
      If .Width > wp(i) Then
        .ShapeRange.PictureFormat.CropRight = (.Width - wp(i)) / sc
      End If
    End With
    Set xlOLE = Nothing
  Else
    objProgCtrl.AddBar lngDivCount:=1
    objProgCtrl.AddBar lngDivCount:=1
  End If
Next
Application.ScreenUpdating = True
DoEvents
apiSleep dwMSec:=500

End Sub


Public Sub TemporaryFolder(ByVal strMode As String)

Dim fsObject As Object

Dim i As Long, j As Long

Set fsObject = CreateObject("Scripting.FileSystemObject")
Select Case strMode
  Case "SetUp"
    lngMC = 0
    tmpFolder = ThisWorkbook.Path & "\Temporary\"
    For i = 1 To UBound(zfdFileName)
      zfdFileName(i) = ""
    Next
    
    With fsObject
      If Not .FolderExists(tmpFolder) Then
        .CreateFolder tmpFolder
      End If
    End With
    
  Case "CleanUp"
    For j = 1 To lngMC
      With rpdApp.Documents
        For i = 0 To .Count - 1
          With .Item(i)
            If .FullName = zfdFileName(j) Then
              .Close
              apiSleep dwMSec:=100
              Exit For
            End If
          End With
        Next
      End With
    Next
    apiSleep dwMSec:=500
    
    For j = 1 To lngMC
      fsObject.DeleteFile zfdFileName(j)
      apiSleep dwMSec:=100
    Next
    apiSleep dwMSec:=500
End Select
Set fsObject = Nothing

End Sub


Public Sub OLEToRapid()

Dim k As Variant

For Each k In ActiveSheet.OLEObjects
  If k.Name = PrimaryOLEObject Or _
     k.Name = SecondaryOLEObject Then
    k.Verb Verb:=xlVerbOpen
    With rpdApp.ActiveDocument
      .FullName = "0"
      .FullName = .DrawInfo.Name & "_" & .DrawInfo.Number
    End With
  End If
Next

End Sub


Public Sub AllSave()

Dim fDialog As FileDialog

Dim i As Long, j As Long
Dim savFolderName As String, savFileCount As String

lngMC = 0
For i = 1 To UBound(zfdFileName)
  zfdFileName(i) = ""
Next

Set fDialog = Application.FileDialog(fileDialogType:= _
                                     msoFileDialogFolderPicker)
With fDialog
  .Title = ProjectTitle
  .ButtonName = "保存先フォルダに指定"
  If .Show = -1 Then
    savFolderName = .SelectedItems(1) & "\"
  Else
    savFolderName = ""
  End If
End With
Set fDialog = Nothing
Application.ScreenUpdating = True

If savFolderName <> "" Then
  With frmMsg
    .Show
    .Image1.Visible = True
    .Label1.Caption = "図面データを保存しています....."
    .Repaint
  End With
  
  With rpdApp.Documents
    For i = 0 To .Count - 1
      With .Item(i)
        If .DrawInfo.Name <> "" And _
           .DrawInfo.Number <> "" Then
          lngMC = lngMC + 1
          zfdFileName(lngMC) = .DrawInfo.Name & "_" & _
                               .DrawInfo.Number & ".zfd"
          .Activate
          .SaveAs sFileName:=savFolderName & zfdFileName(lngMC)
        End If
      End With
    Next
  End With
  
  Unload Object:=frmMsg
  AppActivate Title:=Application.Caption
  
  savFolderName = " 保存先フォルダ : " & vbNewLine & "  " & savFolderName & "    "
  savFileCount = " 保存ファイル数 : " & CStr(lngMC) & "個"
  objError.Message intMsgNumber:=100, _
                   strAddWord1:=savFolderName, _
                   strAddWord2:=savFileCount
End If

End Sub


Public Sub AllPrint()

Dim i As Long

With rpdApp.Documents
  For i = 0 To .Count - 1
    With .Item(i)
      .Activate
      .PrintOut Mode:=zwPrintNormal, bDialog:=True
    End With
  Next
End With

End Sub


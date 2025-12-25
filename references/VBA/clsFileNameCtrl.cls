VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsFileNameCtrl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False


Option Explicit
Option Base 1
Option Compare Binary

Private defName(2) As String, newNames(2, 0 To 5) As String


Public Sub RenameBySummary(ByVal tarFolderName As String)

Dim fsObject As Object

Dim k1 As Variant, k2 As Variant
Dim f As Integer
Dim i As Long, chgCount As Long
Dim logFileName As String, bakFileName As String
Dim tarName As String, chgFileCount As String

Set fsObject = CreateObject("Scripting.FileSystemObject")

logFileName = tarFolderName & "一括リネーム.log"
bakFileName = tarFolderName & "一括リネーム.bak"
If fsObject.FileExists(bakFileName) Then
  fsObject.DeleteFile bakFileName
End If
If fsObject.FileExists(logFileName) Then
  fsObject.GetFile(logFileName).Name = fsObject.GetFileName(bakFileName)
End If

chgCount = 0
If fsObject.GetFolder(tarFolderName).Files.Count > 0 Then
  f = FreeFile(0)
  Open logFileName For Output As #f
  Write #f, "対象フォルダ", tarFolderName
  
  For Each k1 In Worksheets
    If k1.Name <> DefaultSheet1 And _
       k1.Name <> DefaultSheet3 Then
      ReadName tarSheetName:=k1.Name
      For i = 1 To 2
        If defName(i) <> "" And _
           newNames(i, 5) <> "" Then
          For Each k2 In fsObject.GetFolder(tarFolderName).Files
            tarName = fsObject.GetBaseName(k2.Path)
            If Right(tarName, 5) = "(ASP)" Then
              tarName = Mid(tarName, 1, Len(tarName) - 5)
            End If
            If tarName = defName(i) Then
              If fsObject.GetExtensionName(k2.Path) <> "" Then
                newNames(i, 0) = newNames(i, 0) & "_" & newNames(i, 3) & _
                                 "." & fsObject.GetExtensionName(k2.Path)
              End If
              
              If Not fsObject.FileExists(tarFolderName & newNames(i, 0)) Then
                Write #f, k2.Name, newNames(i, 0)
                k2.Name = newNames(i, 0)
                chgCount = chgCount + 1
              End If
              Exit For
            End If
          Next
        End If
      Next
    End If
  Next
  Close #f
End If

If chgCount > 0 Then
  If fsObject.FileExists(bakFileName) Then
    fsObject.DeleteFile bakFileName
  End If
  
  tarFolderName = " 対象フォルダ : " & vbNewLine & "  " & tarFolderName & "    "
  chgFileCount = " リネームファイル数 : " & CStr(chgCount) & "個"
  objError.Message intMsgNumber:=101, _
                   strAddWord1:=tarFolderName, _
                   strAddWord2:=chgFileCount
Else
  If fsObject.FileExists(logFileName) Then
    fsObject.DeleteFile logFileName
  End If
  If fsObject.FileExists(bakFileName) Then
    fsObject.GetFile(bakFileName).Name = fsObject.GetFileName(logFileName)
  End If
  
  objError.Message intMsgNumber:=12
End If

Set fsObject = Nothing

End Sub


Public Sub ListRename()

Dim k As Variant
Dim i As Long, j As Long, filCount As Long
Dim awbName As String
Dim filList(99, 2) As String

awbName = ActiveWorkbook.Name

filCount = 0
For Each k In Worksheets
  If k.Name <> DefaultSheet1 And _
     k.Name <> DefaultSheet3 Then
    ReadName tarSheetName:=k.Name
    For i = 1 To 2
      If defName(i) <> "" Then
        If InStr(k.Name, "ASP") > 0 Then
          defName(i) = defName(i) & "(ASP)"
        End If
        If newNames(i, 5) = "" Then
          newNames(i, 0) = newNames(i, 0) & "<未記入>"
        End If
        filCount = filCount + 1
        filList(filCount, 1) = defName(i)
        filList(filCount, 2) = newNames(i, 0) & "_" & newNames(i, 3)
      End If
    Next
  End If
Next

If filCount > 0 Then
  Workbooks.Open _
    Filename:=ThisWorkbook.Path & "\Template\リネームリスト.xltx"
  DoEvents
  
  Cells(2, 2).Value = Cells(2, 2).Value & " - " & awbName
  For j = 1 To filCount
    Rows(j + 4).Insert _
      Shift:=xlShiftDown, CopyOrigin:=xlFormatFromRightOrBelow
    Cells(j + 4, 2).Formula = _
      "=CONCATENATE(" & _
      Chr(34) & filList(j, 1) & Chr(34) & "," & _
      Chr(34) & "." & Chr(34) & ",$C$" & CStr(j + 7) & ")"
    Cells(j + 4, 3).Value = "⇒"
    Cells(j + 4, 4).Formula = _
      "=CONCATENATE(" & _
      Chr(34) & filList(j, 2) & Chr(34) & "," & _
      Chr(34) & "." & Chr(34) & ",$C$" & CStr(j + 7) & ")"
  Next
End If

End Sub


Public Sub RenameByLogFile(ByVal tarLogFileName As String)

Dim fsObject As Object

Dim k As Variant
Dim f As Integer
Dim chgCount As Long
Dim tarFolderName As String, chgFileCount As String
Dim strOldName As String, strNewName As String

f = FreeFile(0)
Open tarLogFileName For Input As #f
chgCount = 0
tarFolderName = ""
Do While Not EOF(f)
  Input #f, strOldName, strNewName
  If strOldName = "対象フォルダ" Then
    tarFolderName = strNewName
  Else
    Set fsObject = CreateObject("Scripting.FileSystemObject")
    If fsObject.FolderExists(tarFolderName) Then
      If fsObject.GetFolder(tarFolderName).Files.Count > 0 Then
        For Each k In fsObject.GetFolder(tarFolderName).Files
          If k.Name = strNewName Then
            If Not fsObject.FileExists(tarFolderName & strOldName) Then
              k.Name = strOldName
              chgCount = chgCount + 1
            End If
            Exit For
          End If
        Next
      End If
    End If
    Set fsObject = Nothing
  End If
Loop
Close #f

If chgCount > 0 Then
  tarFolderName = " 対象フォルダ : " & vbNewLine & "  " & tarFolderName & "    "
  chgFileCount = " リネームファイル数 : " & CStr(chgCount) & "個"
  objError.Message intMsgNumber:=103, _
                   strAddWord1:=tarFolderName, _
                   strAddWord2:=chgFileCount
Else
  objError.Message intMsgNumber:=13
End If

End Sub


Private Sub ReadName(ByVal tarSheetName As String)

Dim i As Long, j As Long

defName(1) = ""
defName(2) = ""
For j = 0 To UBound(newNames, 2)
  newNames(1, j) = ""
  newNames(2, j) = ""
Next

With Worksheets(tarSheetName)
  newNames(1, 1) = .Cells(2, 3).Value
  newNames(2, 1) = .Cells(2, 3).Value
  newNames(1, 3) = "00"
  newNames(2, 3) = "00"
  If Left(.Name, 1) = "G" And _
     InStr(.Name, "ASP") = 0 Then
    If .Name = "G00" Then
      defName(1) = .Cells(2, 3).Value & "_" & .Name & "_1"
      defName(2) = .Cells(2, 3).Value & "_" & .Name & "_2"
      
      If InStr(.Cells(3, 3).Value, ",") > 0 Then
        newNames(1, 2) = Trim(Left(.Cells(3, 3).Value, _
                             InStr(.Cells(3, 3).Value, ",") - 1))
        newNames(2, 2) = Trim(Mid(.Cells(3, 3).Value, _
                            InStr(.Cells(3, 3).Value, ",") + 1))
        
        newNames(1, 4) = "(" & newNames(1, 2) & ")"
        newNames(2, 4) = "(" & newNames(2, 2) & ")"
        If InStr(.Cells(4, 6).Value, ",") > 0 Then
          newNames(1, 4) = newNames(1, 4) & _
                           Trim(Left(.Cells(4, 6).Value, _
                               InStr(.Cells(4, 6).Value, ",") - 1))
          newNames(2, 4) = newNames(2, 4) & _
                           Trim(Mid(.Cells(4, 6).Value, _
                              InStr(.Cells(4, 6).Value, ",") + 1))
        End If
        
        If Not IsEmpty(.Cells(5, 3).Value) Then
          newNames(1, 5) = Left(Trim(.Cells(5, 3).Value), 13)
          If Len(Trim(.Cells(5, 3).Value)) > 14 Then
            newNames(2, 5) = Right(Trim(.Cells(5, 3).Value), 13)
          End If
        End If
      Else
        newNames(1, 2) = Trim(.Cells(3, 3).Value) & "(1_2)"
        newNames(2, 2) = Trim(.Cells(3, 3).Value) & "(2_2)"
        
        newNames(1, 4) = "(" & Trim(.Cells(3, 3).Value) & ")"
        newNames(2, 4) = "(" & Trim(.Cells(3, 3).Value) & ")"
        If InStr(.Cells(4, 6).Value, ",") > 0 Then
          newNames(1, 4) = newNames(1, 4) & _
                           Trim(Left(.Cells(4, 6).Value, _
                               InStr(.Cells(4, 6).Value, ",") - 1))
          newNames(2, 4) = newNames(2, 4) & _
                           Trim(Mid(.Cells(4, 6).Value, _
                              InStr(.Cells(4, 6).Value, ",") + 1))
        Else
          newNames(1, 4) = newNames(1, 4) & "(1_2)"
          newNames(2, 4) = newNames(2, 4) & "(2_2)"
        End If
        
        If Not IsEmpty(.Cells(5, 3).Value) Then
          newNames(1, 5) = Left(Trim(.Cells(5, 3).Value), 13)
          newNames(2, 5) = Left(Trim(.Cells(5, 3).Value), 13)
        End If
      End If
    Else
      defName(1) = .Cells(2, 3).Value & "_" & .Name
      defName(2) = .Cells(2, 3).Value & "_" & "Z" & Mid(.Name, 2, 2)
      
      newNames(1, 2) = Trim(.Cells(3, 3).Value)
      newNames(2, 2) = "Z" & Mid(Trim(.Cells(3, 3).Value), 2, 2)
      
      newNames(1, 4) = "(" & Trim(.Cells(3, 3).Value) & ")"
      newNames(2, 4) = "(Z" & Mid(Trim(.Cells(3, 3).Value), 2, 2) & ")プレス図"
      
      If Not IsEmpty(.Cells(5, 3).Value) Then
        newNames(1, 5) = Left(Trim(.Cells(5, 3).Value), 13)
        If Len(Trim(.Cells(5, 3).Value)) > 14 Then
          newNames(2, 5) = Right(Trim(.Cells(5, 3).Value), 13)
        End If
      End If
    End If
  Else
    defName(1) = .Cells(2, 3).Value & "_" & Left(.Name, 3)
    
    If Left(.Name, 1) = "G" And _
       InStr(.Name, "ASP") > 0 Then
      newNames(1, 2) = Trim(.Cells(3, 3).Value)
      newNames(1, 4) = "(" & Trim(.Cells(3, 3).Value) & ")ＧＭ"
    ElseIf Left(.Name, 1) = "B" Then
      newNames(1, 2) = Left(Trim(.Cells(3, 3).Value), 3)
      newNames(1, 4) = "(" & Left(Trim(.Cells(3, 3).Value), 3) & ")"
      If InStr(.Cells(3, 3).Value, "(") > 0 And _
         InStr(.Cells(3, 3).Value, ")") > 0 Then
        newNames(1, 4) = newNames(1, 4) & _
                         Trim(Mid(.Cells(3, 3).Value, InStr(.Cells(3, 3).Value, "(") + 1, _
                         InStr(.Cells(3, 3).Value, ")") - (InStr(.Cells(3, 3).Value, "(") + 1)))
      Else
        newNames(1, 4) = newNames(1, 4) & Trim(Mid(.Cells(3, 3).Value, 4))
      End If
    Else
      newNames(1, 2) = Trim(.Cells(3, 3).Value)
      newNames(1, 4) = "(" & Trim(.Cells(3, 3).Value) & ")"
    End If
    
    If Not IsEmpty(.Cells(5, 3).Value) Then
      newNames(1, 5) = Left(Trim(.Cells(5, 3).Value), 13)
    End If
  End If
  
  For i = 1 To 2
    If defName(i) <> "" Then
      For j = 1 To UBound(newNames, 2)
        newNames(i, j) = Replace(newNames(i, j), "/", "_")
        newNames(i, j) = Replace(newNames(i, j), " ", "")
        Select Case j
          Case 1: newNames(i, 0) = newNames(i, 1)
          Case 5: newNames(i, 0) = newNames(i, 0) & "_" & newNames(i, 5)
          Case Else: newNames(i, 0) = newNames(i, 0) & "-" & newNames(i, j)
        End Select
      Next
    End If
  Next
End With

End Sub


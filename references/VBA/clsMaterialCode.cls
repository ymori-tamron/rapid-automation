VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsMaterialCode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False


Option Explicit
Option Base 1
Option Compare Binary

Private blnOutStt As Boolean

Private strCode(99, 3) As String


Private Sub Class_Initialize()

Dim i As Long, j As Long

blnOutStt = False
For j = 1 To 3
  For i = 1 To UBound(strCode, 1)
    strCode(i, j) = ""
  Next
Next

End Sub


Public Property Get blnStatus() As Boolean

blnStatus = blnOutStt

End Property


Public Sub ReadCode()

Dim FileToOpen As Variant, Filter As Variant
Dim k As Variant
Dim f As Integer
Dim i As Long, maxRow As Long
Dim sepStr As String, spcStr As String
Dim TextLine As String, txt As String

Filter = "ÉeÉLÉXÉgÉtÉ@ÉCÉã (*.txt;*.csv),*.txt;*.csv," & _
         "Ç∑Ç◊ÇƒÇÃÉtÉ@ÉCÉã (*.*),*.*"
FileToOpen = Application.GetOpenFilename(FileFilter:=Filter, _
                                         FilterIndex:=1, _
                                         Title:=ProjectTitle)
If FileToOpen <> False Then
  blnOutStt = False
  
  Select Case UCase(Right(FileToOpen, 3))
    Case "TXT": sepStr = vbTab
    Case "CSV": sepStr = ","
    Case Else: sepStr = ""
  End Select
  If sepStr <> "" Then
    spcStr = sepStr
    For Each k In Worksheets
      If k.Name = "G00" Then
        If InStr(k.Cells(3, 3).Value, ",") > 0 Then
          spcStr = Trim(Mid(k.Cells(3, 3).Value, _
                      InStr(k.Cells(3, 3).Value, ",") + 1))
        End If
        Exit For
      End If
    Next
    
    f = FreeFile(0)
    Open FileToOpen For Input As #f
    maxRow = 0
    Do While Not EOF(f)
      Line Input #f, TextLine
      If InStr(TextLine, sepStr) > 0 Then
        If Right(TextLine, 1) <> sepStr Then TextLine = TextLine & sepStr
        
        maxRow = maxRow + 1
        If maxRow > UBound(strCode, 1) Then Exit Do
        
        strCode(maxRow, 1) = ""
        strCode(maxRow, 2) = ""
        strCode(maxRow, 3) = ""
        txt = ""
        For i = 1 To Len(TextLine)
          If Mid(TextLine, i, 1) <> sepStr Then
            If Mid(TextLine, i, 1) <> Chr(34) And _
               Mid(TextLine, i, 1) <> Chr(39) Then
              txt = txt & Mid(TextLine, i, 1)
            End If
          ElseIf txt <> "" Then
            txt = Trim(txt)
            If txt Like "*-0000*" Then
              If Len(txt) >= 13 Then
                If strCode(maxRow, 1) = "" Then
                  strCode(maxRow, 1) = Left(txt, 13)
                End If
              End If
            ElseIf txt Like "G##*" Or _
                   txt Like "B##*" Or _
                   txt Like "Z##*" Or _
                   txt Like (spcStr & "*") Then
              If strCode(maxRow, 2) = "" Then
                strCode(maxRow, 2) = StrConv(txt, vbNarrow)
              End If
            ElseIf txt Like "ASSY*" Or _
                   txt Like "ÉåÉìÉY*" Then
              strCode(maxRow, 3) = StrConv(txt, vbWide)
            End If
            txt = ""
          End If
        Next
        
        If strCode(maxRow, 1) = "" Or _
           strCode(maxRow, 2) = "" Or _
           strCode(maxRow, 3) = "" Then
          strCode(maxRow, 1) = ""
          strCode(maxRow, 2) = ""
          strCode(maxRow, 3) = ""
          maxRow = maxRow - 1
        Else
          blnOutStt = True
        End If
      End If
    Loop
    Close #f
  End If
  
  If Not blnOutStt Then
    objError.Message intMsgNumber:=6
  End If
End If

End Sub


Public Sub WriteCode()

Dim i As Long
Dim asName As String
Dim txt As String, tmp As String

asName = ActiveSheet.Name

Select Case Left(asName, 1)
  Case "G"
    For i = 1 To UBound(strCode, 1)
      If Left(strCode(i, 2), 3) = Left(asName, 3) Then
        If strCode(i, 3) = "Ç`ÇrÇrÇxÅ^ÉtÉ@ÉìÉgÉÄ" Or _
           strCode(i, 3) = "ÉåÉìÉYÅ^ãÖñ " Or _
           strCode(i, 3) = "ÉåÉìÉYÅ^ÇfÇl" Or _
           strCode(i, 3) = "ÉåÉìÉYÅ^ÉvÉâÉåÉìÉY" Or _
           strCode(i, 3) = "ÉåÉìÉYÅ^ÇªÇÃëº" Then
          If Len(Trim(Cells(5, 3).Value)) > 14 Then
            tmp = "," & Right(Trim(Cells(5, 3).Value), 13)
          Else
            tmp = ""
          End If
          Cells(5, 3).Value = strCode(i, 1) & tmp
          Exit For
        End If
      End If
    Next
    
    If asName = "G00" Then
      If InStr(Cells(3, 3).Value, ",") > 0 Then
        txt = Trim(Mid(Cells(3, 3).Value, _
                 InStr(Cells(3, 3).Value, ",") + 1))
        For i = 1 To UBound(strCode, 1)
          If Left(strCode(i, 2), Len(txt)) = txt Then
            If strCode(i, 3) = "Ç`ÇrÇrÇxÅ^ÉtÉ@ÉìÉgÉÄ" Or _
               strCode(i, 3) = "ÉåÉìÉYÅ^ÇªÇÃëº" Then
              If Len(Trim(Cells(5, 3).Value)) > 12 Then
                tmp = Left(Trim(Cells(5, 3).Value), 13)
              Else
                tmp = String(13, "x")
              End If
              Cells(5, 3).Value = tmp & "," & strCode(i, 1)
              Exit For
            End If
          End If
        Next
      End If
    ElseIf InStr(asName, "ASP") = 0 Then
      For i = 1 To UBound(strCode, 1)
        If Left(strCode(i, 2), 3) = "Z" & Mid(asName, 2, 2) Then
          If strCode(i, 3) = "ÉåÉìÉYÅ^è…çﬁ" Then
            If Len(Trim(Cells(5, 3).Value)) > 12 Then
              tmp = Left(Trim(Cells(5, 3).Value), 13)
            Else
              tmp = String(13, "x")
            End If
            Cells(5, 3).Value = tmp & "," & strCode(i, 1)
            Exit For
          End If
        End If
      Next
    End If
    
  Case "B"
    For i = 1 To UBound(strCode, 1)
      If Left(strCode(i, 2), 3) = asName Then
        If strCode(i, 3) = "ÉåÉìÉYÅ^ê⁄çá" Then
          Cells(5, 3).Value = strCode(i, 1)
          Exit For
        End If
      End If
    Next
    
  Case "H"
    For i = 1 To UBound(strCode, 1)
      If Left(strCode(i, 2), 3) = "G" & Mid(asName, 2, 2) Then
        If strCode(i, 3) = "ÉåÉìÉYÅ^ï°çáïî" Then
          Cells(5, 3).Value = strCode(i, 1)
          Exit For
        End If
      End If
    Next
    
  Case "M"
    For i = 1 To UBound(strCode, 1)
      If Left(strCode(i, 2), 3) = "G" & Mid(asName, 2, 2) Then
        If strCode(i, 3) = "ÉåÉìÉYÅ^ï°çá" Then
          Cells(5, 3).Value = strCode(i, 1)
          Exit For
        End If
      End If
    Next
End Select

End Sub


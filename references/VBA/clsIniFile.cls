VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsIniFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False


Option Explicit
Option Base 1
Option Compare Binary

Private iniFileName As String


Private Sub Class_Initialize()

iniFileName = ThisWorkbook.Path & "\" & ProjectTitle & ".ini"

End Sub


Public Sub FileToArray()

Dim f As Integer
Dim i As Long, tmp As Long
Dim lngIndex As Long
Dim strValue As String

For i = 1 To UBound(Alpha)
  Alpha(i) = "0"
Next
For i = 1 To UBound(fCheck)
  fCheck(i) = "False"
Next
For i = 1 To UBound(vCheck)
  vCheck(i) = "False"
Next

f = FreeFile(0)
Open iniFileName For Input As #f
For i = 1 To UBound(Alpha)
  If Not EOF(f) Then
    Input #f, lngIndex, strValue
    If lngIndex >= 1 And _
       lngIndex <= UBound(Alpha) Then
      Alpha(lngIndex) = strValue
    End If
  End If
Next

For i = 1 To UBound(fCheck)
  If Not EOF(f) Then
    Input #f, lngIndex, strValue
    tmp = lngIndex - UBound(Alpha)
    If tmp >= 1 And _
       tmp <= UBound(fCheck) Then
      fCheck(tmp) = strValue
    End If
  End If
Next

For i = 1 To UBound(vCheck)
  If Not EOF(f) Then
    Input #f, lngIndex, strValue
    tmp = lngIndex - (UBound(Alpha) + UBound(fCheck))
    If tmp >= 1 And _
       tmp <= UBound(vCheck) Then
      vCheck(tmp) = strValue
    End If
  End If
Next
Close #f

End Sub


Public Sub FileToForm()

Dim k As Variant
Dim f As Integer
Dim i As Long, tmp As Long
Dim lngIndex As Long
Dim strValue As String

For Each k In frmIni.Controls
  Select Case TypeName(k)
    Case "TextBox": k.Value = "NaN"
    Case "CheckBox": k.Value = False
  End Select
Next

f = FreeFile(0)
Open iniFileName For Input As #f
For i = 1 To UBound(Alpha)
  If Not EOF(f) Then
    Input #f, lngIndex, strValue
    If lngIndex >= 1 And _
       lngIndex <= UBound(Alpha) Then
      frmIni.Controls("TextBox" & CStr(lngIndex)).Value = strValue
    End If
  End If
Next

For i = 1 To UBound(fCheck)
  If Not EOF(f) Then
    Input #f, lngIndex, strValue
    tmp = lngIndex - UBound(Alpha)
    If tmp >= 1 And _
       tmp <= UBound(fCheck) Then
      frmIni.Controls("CheckBoxF" & CStr(tmp)).Value = CBool(strValue)
    End If
  End If
Next

For i = 1 To UBound(vCheck)
  If Not EOF(f) Then
    Input #f, lngIndex, strValue
    tmp = lngIndex - (UBound(Alpha) + UBound(fCheck))
    If tmp >= 1 And _
       tmp <= UBound(vCheck) Then
      frmIni.Controls("CheckBoxV" & CStr(tmp)).Value = CBool(strValue)
    End If
  End If
Next
Close #f

End Sub


Public Sub FormToFile()

Dim f As Integer
Dim i As Long

f = FreeFile(0)
Open iniFileName For Output As #f
For i = 1 To UBound(Alpha)
  Write #f, i, frmIni.Controls("TextBox" & CStr(i)).Value
Next

For i = 1 To UBound(fCheck)
  Write #f, i + UBound(Alpha), _
            CStr(frmIni.Controls("CheckBoxF" & CStr(i)).Value)
Next

For i = 1 To UBound(vCheck)
  Write #f, i + (UBound(Alpha) + UBound(fCheck)), _
            CStr(frmIni.Controls("CheckBoxV" & CStr(i)).Value)
Next
Close #f
apiSleep dwMSec:=500

End Sub


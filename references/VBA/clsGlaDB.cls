VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsGlaDB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False


Option Explicit
Option Base 1
Option Compare Binary

Private xlWbOutSum As Excel.Workbook
Private xlWbOutGDB As Excel.Workbook

Private strOutWsn As String


Private Sub Class_Initialize()

Set xlWbOutSum = Nothing
Set xlWbOutGDB = Nothing

strOutWsn = ""

End Sub


Private Sub Class_Terminate()

Set xlWbOutSum = Nothing
Set xlWbOutGDB = Nothing

strOutWsn = ""

objRibbonCtrl.SetEnabled lngIndex:=5, blnValue:=False
objRibbonCtrl.SetEnabled lngIndex:=6, blnValue:=False

End Sub


Public Property Get xlWb_DataSummary() As Excel.Workbook

Set xlWb_DataSummary = xlWbOutSum

End Property


Public Property Get xlWb_GlaDB() As Excel.Workbook

Set xlWb_GlaDB = xlWbOutGDB

End Property


Public Property Get strWsName() As String

strWsName = strOutWsn

End Property


Public Sub Load()

Dim k As Variant
Dim chkWs As Integer
Dim i As Long, j As Long, cc As Long, maxCol As Long
Dim dblGla(4) As Double
Dim tmp As String, strGla(2) As String

strGla(1) = ""
strGla(2) = ""
ClearSumInfo
If Workbooks.Count > 0 Then
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
    If ActiveSheet.Name <> "G00" Then
      If Left(ActiveSheet.Name, 1) = "G" Then
        Set xlWbOutSum = ActiveWorkbook
        strOutWsn = ActiveSheet.Name
        objRibbonCtrl.SetEnabled lngIndex:=5, blnValue:=True
        
        strGla(1) = Cells(11, 3).Value
        strGla(2) = Cells(12, 3).Value
        dblGla(1) = Cells(8, 9).Value
        dblGla(2) = Cells(9, 9).Value
        With Worksheets(DefaultSheet3)
          cc = 0
          maxCol = .Cells(6, .Columns.Count).End(xlToLeft).Column
          If maxCol > 2 Then
            For i = 3 To maxCol Step 3
              If .Cells(6, i).Value = ActiveSheet.Name Then
                cc = i: Exit For
              End If
            Next
          End If
          
          If cc > 0 Then
            dblGla(3) = .Cells(12, cc).Value
            dblGla(4) = .Cells(13, cc).Value
          End If
        End With
      End If
    End If
  End If
End If

If xlWbOutGDB Is Nothing Then
  Workbooks.Open _
    Filename:=ThisWorkbook.Path & "\Utility\Glasses-DataBase.xlsm"
  DoEvents
  Set xlWbOutGDB = ActiveWorkbook
  objRibbonCtrl.SetEnabled lngIndex:=6, blnValue:=True
Else
  xlWbOutGDB.Activate
End If

If strGla(1) <> "" And _
   strGla(2) <> "" Then
  With Worksheets("ŒŸõ")
    With .OLEObjects("Combo_manu").Object
      If .ListCount = 0 Then
        .AddItem ""
        .AddItem "Ohara"
        .AddItem "Hikari"
        .AddItem "HOYA"
        .AddItem "CDGM"
        .AddItem "SCHOTT"
        .AddItem "V‰ØŒõ"
      End If
      
      For i = 0 To .ListCount - 1
        If UCase(.List(i, 0)) = strGla(2) Then
          .ListIndex = i
          Exit For
        End If
      Next
    End With
    
    With .OLEObjects("Combo_glass").Object
      If .ListCount > 0 Then
        For i = 0 To .ListCount - 1
          If InStr(.List(i, 0), " ") > 0 Then
            tmp = ""
            For j = 1 To Len(.List(i, 0))
              If Mid(.List(i, 0), j, 1) <> " " Then
                tmp = tmp & Mid(.List(i, 0), j, 1)
              End If
            Next
          Else
            tmp = .List(i, 0)
          End If
          
          If UCase(tmp) = strGla(1) Then
            .ListIndex = i
            Exit For
          End If
        Next
      End If
    End With
    
    .Cells(2, 2).Value = dblGla(1)
    .Cells(4, 2).Value = dblGla(2)
    .Cells(2, 3).Value = dblGla(3)
    .Cells(2, 4).Value = dblGla(3)
    .Cells(4, 3).Value = dblGla(4) * 100
    .Cells(4, 4).Value = dblGla(4) * 100
    
    .Range(.Rows(9), .Rows(25)).ClearContents
    .Range(.Rows(28), .Rows(200)).ClearContents
  End With
End If

End Sub


Public Sub PickUp()

Dim dicObject As Object

Dim i As Long, j As Long
Dim tmp As String

Set dicObject = CreateObject("Scripting.Dictionary")
With xlWbOutGDB.Worksheets("ŒŸõ")
  For i = 9 To 25
    If Not IsEmpty(.Cells(i, 1).Value) And _
       Not IsEmpty(.Cells(i, 2).Value) Then
      If Not dicObject.Exists(UCase(.Cells(i, 1).Value)) Then
        If InStr(.Cells(i, 2).Value, " ") > 0 Then
          tmp = ""
          For j = 1 To Len(.Cells(i, 2).Value)
            If Mid(.Cells(i, 2).Value, j, 1) <> " " Then
              tmp = tmp & Mid(.Cells(i, 2).Value, j, 1)
            End If
          Next
        Else
          tmp = .Cells(i, 2).Value
        End If
        
        dicObject.Add UCase(.Cells(i, 1).Value), UCase(tmp)
      End If
    End If
  Next
End With

xlWbOutSum.Activate
Worksheets(strOutWsn).Activate
For i = 31 To 36
  If dicObject.Exists(Cells(i, 5).Value) Then
    Cells(i, 6).Value = dicObject.Item(Cells(i, 5).Value)
  End If
Next
If dicObject.Exists("HIKARI") Then
  Cells(34, 6).Value = dicObject.Item("HIKARI")
End If
If dicObject.Exists("CDGM") Then
  Cells(35, 6).Value = dicObject.Item("CDGM")
End If

ClearSumInfo
Set dicObject = Nothing

End Sub


Public Sub ClearSumInfo()

Set xlWbOutSum = Nothing
strOutWsn = ""
objRibbonCtrl.SetEnabled lngIndex:=5, blnValue:=False

End Sub


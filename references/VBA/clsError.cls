VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsError"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False


Option Explicit
Option Base 1
Option Compare Binary

Private intOut As Integer, intRtn As Integer


Public Property Get intCode() As Integer

intCode = intOut

End Property


Public Property Get intReturnValue() As Integer

intReturnValue = intRtn

End Property


Public Sub CheckConvert()

Dim k As Variant
Dim i As Long

If Workbooks.Count > 0 Then
  For Each k In Workbooks
    If k.BuiltinDocumentProperties("Creation Date").Value = curWb Then
      k.Activate
      If Worksheets.Count <> 3 Then
        intOut = 3
      ElseIf Worksheets(1).Name <> DefaultSheet1 Or _
             Worksheets(2).Name <> DefaultSheet2 Or _
             Worksheets(3).Name <> DefaultSheet3 Then
        intOut = 4
      Else
        intOut = 0
      End If
      Exit For
    Else
      intOut = 2
    End If
  Next
Else
  intOut = 1
End If

Select Case intOut
  Case 1, 2, 3
    Message intMsgNumber:=intOut
    Reset
    Exit Sub
  Case 4
    Message intMsgNumber:=intOut
    Exit Sub
End Select

If lngSur > 3 Then
  For i = 7 To lngSur + 3
    If Not IsEmpty(Cells(i, 3).Value) And _
       Cells(i, 3).Value <> "STO" And _
       Left(Cells(i, 3).Value, 2) <> "AP" Then
      If IsEmpty(Cells(i, 9).Value) Or _
         IsEmpty(Cells(i + 1, 9).Value) Then
        Message intMsgNumber:=200
        Select Case intRtn
          Case vbOK
            objRibbonCtrl.SetEnabled lngIndex:=1, blnValue:=False
            objRibbonCtrl.SetEnabled lngIndex:=2, blnValue:=True
          Case vbCancel
            Reset
        End Select
        intOut = 200: Exit Sub
      End If
    End If
  Next
End If

End Sub


Public Sub CheckSummary()

Dim k As Variant

If Workbooks.Count > 0 Then
  intOut = -2
  If Worksheets.Count > 2 Then
    For Each k In Worksheets
      If k.Name = DefaultSheet1 Or _
         k.Name = DefaultSheet3 Then
        intOut = intOut + 1
      Else
        If k.Name = DefaultSheet2 Then
          intOut = 9: Exit For
        ElseIf Left(k.Name, 1) <> "G" And _
               Left(k.Name, 1) <> "B" And _
               Left(k.Name, 1) <> "H" And _
               Left(k.Name, 1) <> "M" Then
          intOut = 9: Exit For
        ElseIf Not IsNumeric(Mid(k.Name, 2, 2)) Then
          intOut = 9: Exit For
        End If
      End If
    Next
  Else
    intOut = 8
  End If
Else
  intOut = 1
End If

Select Case intOut
  Case 1, 8, 9
    Message intMsgNumber:=intOut
  Case Is < 0
    Message intMsgNumber:=4
End Select

End Sub


Public Sub CheckRapidExists()

With rpdApp.Documents
  If .Count > 0 Then
    intOut = 0
  Else
    intOut = 20
  End If
End With

If intOut = 20 Then
  AppActivate Title:=Application.Caption
  Message intMsgNumber:=intOut
End If

End Sub


Public Sub CheckRapidMatch()

Dim k As Variant
Dim i As Long
Dim prtNum As String

CheckSummary
If intOut <> 0 Then Exit Sub

CheckRapidExists
If intOut <> 0 Then Exit Sub

With rpdApp.Documents
  intOut = 21
  If ActiveSheet.Name <> DefaultSheet1 And _
     ActiveSheet.Name <> DefaultSheet3 Then
    mdlNum = Cells(2, 3).Value
    prtNum = ActiveSheet.Name
    For i = 0 To .Count - 1
      With .Item(i).DrawInfo
        If .Name = mdlNum Then
          If .Number = prtNum & "_1" Or _
             prtNum <> "G00" And _
             Left(.Number, 3) = Left(prtNum, 3) Then
            intOut = 0: Exit Sub
          End If
          
          If Left(prtNum, 1) = "G" Then
            If .Number = prtNum & "_2" Or _
               .Number = "Z" & Mid(prtNum, 2, 2) Then
              intOut = 0: Exit Sub
            End If
          End If
        End If
      End With
    Next
  Else
    For Each k In Worksheets
      If k.Name <> DefaultSheet1 And _
         k.Name <> DefaultSheet3 Then
        mdlNum = k.Cells(2, 3).Value
        prtNum = k.Name
        For i = 0 To .Count - 1
          With .Item(i).DrawInfo
            If .Name = mdlNum Then
              If .Number = prtNum & "_1" Or _
                 prtNum <> "G00" And _
                 Left(.Number, 3) = Left(prtNum, 3) Then
                intOut = 0: Exit Sub
              End If
              
              If Left(prtNum, 1) = "G" Then
                If .Number = prtNum & "_2" Or _
                   .Number = "Z" & Mid(prtNum, 2, 2) Then
                  intOut = 0: Exit Sub
                End If
              End If
            End If
          End With
        Next
      End If
    Next
  End If
End With

If intOut = 21 Then
  AppActivate Title:=Application.Caption
  Message intMsgNumber:=intOut
End If

End Sub


Public Sub CheckOLE()

Dim k1 As Variant, k2 As Variant

CheckSummary
If intOut <> 0 Then Exit Sub

intOut = 22
If ActiveSheet.Name <> DefaultSheet1 And _
   ActiveSheet.Name <> DefaultSheet3 Then
  If ActiveSheet.OLEObjects.Count > 0 Then
    For Each k2 In ActiveSheet.OLEObjects
      If k2.Name = PrimaryOLEObject Or _
         k2.Name = SecondaryOLEObject Then
        intOut = 0: Exit Sub
      End If
    Next
  End If
Else
  For Each k1 In Worksheets
    If k1.Name <> DefaultSheet1 And _
       k1.Name <> DefaultSheet3 Then
      If k1.OLEObjects.Count > 0 Then
        For Each k2 In k1.OLEObjects
          If k2.Name = PrimaryOLEObject Or _
             k2.Name = SecondaryOLEObject Then
            intOut = 0: Exit Sub
          End If
        Next
      End If
    End If
  Next
End If

If intOut = 22 Then
  Message intMsgNumber:=intOut
End If

End Sub


Public Sub CheckRapidFName()

Dim i As Long

CheckRapidExists
If intOut <> 0 Then Exit Sub

With rpdApp.Documents
  intOut = 23
  For i = 0 To .Count - 1
    With .Item(i).DrawInfo
      If .Name <> "" And _
         .Number <> "" Then
        intOut = 0: Exit Sub
      End If
    End With
  Next
End With

If intOut = 23 Then
  AppActivate Title:=Application.Caption
  Message intMsgNumber:=intOut
End If

End Sub


Public Sub CheckCFG()

Dim i As Long
Dim defName(9) As String

defName(1) = "TxtA4"
defName(2) = "TxtA4Assy"
defName(3) = "TxtA4Press"
defName(4) = "TxtA4Balsam"
defName(5) = "TxtA3"
defName(6) = "TxtA3E"
defName(7) = "TxtA3Assy"
defName(8) = "Param"
defName(9) = "PressParam"

If Workbooks.Count > 0 Then
  intOut = 10
  For i = 1 To UBound(defName)
    If ActiveSheet.Name = defName(i) Then
      intOut = 0: Exit For
    End If
  Next
Else
  intOut = 1
End If

Select Case intOut
  Case 1, 10
    Message intMsgNumber:=intOut
End Select

End Sub


Public Sub CheckManualExists()

Dim fsObject As Object

Dim mnlFileName As String

Set fsObject = CreateObject("Scripting.FileSystemObject")
mnlFileName = ThisWorkbook.Path & "\取扱説明書.pdf"
If fsObject.FileExists(mnlFileName) Then
  intOut = 0
Else
  intOut = 15
End If
Set fsObject = Nothing

If intOut = 15 Then
  Message intMsgNumber:=intOut, _
          strAddWord1:=mnlFileName
End If

End Sub


Public Sub CheckUtilityExists()

Dim fsObject As Object

Dim utlFileName As String

Set fsObject = CreateObject("Scripting.FileSystemObject")
utlFileName = ThisWorkbook.Path & "\Utility\Utility.xlam"
If fsObject.FileExists(utlFileName) Then
  intOut = 0
Else
  intOut = 16
End If
Set fsObject = Nothing

If intOut = 16 Then
  Message intMsgNumber:=intOut, _
          strAddWord1:=utlFileName
End If

End Sub


Public Sub CheckGlaDBExists()

Dim fsObject As Object

Dim gdbFileName As String

Set fsObject = CreateObject("Scripting.FileSystemObject")
gdbFileName = ThisWorkbook.Path & "\Utility\Glasses-DataBase.xlsm"
If fsObject.FileExists(gdbFileName) Then
  intOut = 0
Else
  intOut = 30
End If
Set fsObject = Nothing

If intOut = 30 Then
  Message intMsgNumber:=intOut, _
          strAddWord1:=gdbFileName
End If

End Sub


Public Sub CheckGlaDBStatus()

Dim dicObject As Object

Dim k As Variant
Dim i As Long

If Workbooks.Count > 0 Then
  If Not objGlaDB Is Nothing Then
    With objGlaDB
      intOut = -2
      If Not .xlWb_DataSummary Is Nothing And _
         Not .xlWb_GlaDB Is Nothing Then
        If .xlWb_DataSummary.Worksheets.Count > 2 And _
           .xlWb_GlaDB.Worksheets.Count > 0 Then
          For Each k In .xlWb_DataSummary.Worksheets
            If k.Name = .strWsName Then
              intOut = intOut + 1: Exit For
            End If
          Next
          
          For Each k In .xlWb_GlaDB.Worksheets
            If k.Name = "検索" Then
              intOut = intOut + 1: Exit For
            End If
          Next
        Else
          intOut = 8
        End If
      End If
    End With
  Else
    intOut = 7
  End If
Else
  intOut = 1
End If

Select Case intOut
  Case 1, 7, 8
    Message intMsgNumber:=intOut
    Exit Sub
  Case Is < 0
    Message intMsgNumber:=31
    Exit Sub
End Select

Set dicObject = CreateObject("Scripting.Dictionary")
With objGlaDB.xlWb_GlaDB.Worksheets("検索")
  For i = 9 To 25
    If Not IsEmpty(.Cells(i, 1).Value) And _
       Not IsEmpty(.Cells(i, 2).Value) Then
      If Not dicObject.Exists(UCase(.Cells(i, 1).Value)) Then
        dicObject.Add UCase(.Cells(i, 1).Value), ""
      Else
        intOut = 32: Exit For
      End If
    End If
  Next
End With
Set dicObject = Nothing

If intOut = 32 Then
  Message intMsgNumber:=intOut
End If

End Sub


Public Sub Message(ByVal intMsgNumber As Integer, _
                   Optional ByVal strAddWord1 As String = "", _
                   Optional ByVal strAddWord2 As String = "", _
                   Optional ByVal strAddWord3 As String = "")

Dim btn As Long
Dim msg As String

If intMsgNumber < 100 Then
  btn = vbOKOnly + vbExclamation + vbDefaultButton1
  msg = "ERROR : " & CStr(intMsgNumber) & vbNewLine & vbNewLine
ElseIf intMsgNumber < 200 Then
  btn = vbOKOnly + vbInformation + vbDefaultButton1
Else
  btn = vbOKCancel + vbQuestion + vbDefaultButton1
End If

Select Case intMsgNumber
  Case -1
    msg = msg & _
          "この機能は、現在開発中です    "
  Case 0
    msg = msg & _
          "無効なコマンドです    "
  Case 1
    msg = msg & _
          "ワークブックが存在しません" & vbNewLine & _
          "作業を始めからやり直して下さい    "
  Case 2
    msg = msg & _
          "読込中のワークブックが見当たりません    " & vbNewLine & _
          "作業を始めからやり直して下さい"
  Case 3
    msg = msg & _
          "デフォルトのシート配置ではありません    " & vbNewLine & _
          "作業を始めからやり直して下さい"
  Case 4
    msg = msg & _
          "デフォルトのシートが見当たりません" & vbNewLine & _
          "シート名を変更した場合は、元に戻して下さい    "
  Case 5
    msg = msg & _
          "図面データサマリとして認識出来ません    "
  Case 6
    msg = msg & _
          "品目コードの読込に失敗しました" & vbNewLine & _
          "所定の内容、対応する形式のファイルを指定して下さい    "
  Case 7
    msg = msg & _
          "処理に必要なメモリ情報が消失しています    " & vbNewLine & _
          "作業を始めからやり直して下さい"
  Case 8
    msg = msg & _
          "処理に必要な情報が不足しています" & vbNewLine & _
          "シートを移動･削除した場合は、元に戻して下さい    "
  Case 9
    msg = msg & _
          "シート名にデフォルトの部番以外が指定されています" & vbNewLine & _
          "シート名を変更した場合は、元に戻して下さい" & vbNewLine & _
          vbNewLine & _
          " NOTE : 複合非球面関連の'M??'や'H??'は、デフォルトです    "
  Case 10
    msg = msg & _
          "対応するCFGファイルが見当たりません" & vbNewLine & _
          "シート名を変更した場合は、元に戻して下さい    "
  Case 11
    msg = msg & _
          "CFGファイル内にテキスト座標データが見当たりません    " & vbNewLine & _
          "'Txt･･･'で始まるCFGファイルを指定して下さい"
  Case 12
    msg = msg & _
          "リネーム処理に失敗しました" & vbNewLine & _
          vbNewLine & _
          "図面データサマリと対応が取れない名前のファイルや    " & vbNewLine & _
          "対応するシートの品目コード欄が未記入のものは、" & vbNewLine & _
          "リネーム処理が出来ません"
  Case 13
    msg = msg & _
          "ログファイルの内容と一致するファイルが見当たらない為、    " & vbNewLine & _
          "ファイル名を元に戻せませんでした"
  Case 14
    msg = msg & _
          "いずれのシートにも修正の必要な箇所が見当たりません    "
  Case 15
    msg = msg & _
          "'取扱説明書'のファイルが存在しません" & vbNewLine & _
          vbNewLine & _
          "'" & strAddWord1 & "'    " & vbNewLine & _
          vbNewLine & _
          "上記のファイルを移動･削除したり、ファイル名や    " & vbNewLine & _
          "拡張子を変更した場合は、元に戻して下さい"
  Case 16
    msg = msg & _
          "'Utility'のファイルが存在しません" & vbNewLine & _
          vbNewLine & _
          "'" & strAddWord1 & "'    " & vbNewLine & _
          vbNewLine & _
          "上記のファイルを移動･削除したり、ファイル名や    " & vbNewLine & _
          "拡張子を変更した場合は、元に戻して下さい"
  Case 20
    msg = msg & _
          "図脳RAPIDで開いている図面データが存在しません    "
  Case 21
    msg = msg & _
          "対応する図面データが見当たりません    "
  Case 22
    msg = msg & _
          "読込済の図面データが見当たりません    "
  Case 23
    msg = msg & _
          "保存可能な図面データが見当たりません" & vbNewLine & _
          "図面情報の図面名と図面番号を指定して下さい    "
  Case 30
    msg = msg & _
          "'類似硝種検索ツール'のファイルが存在しません" & vbNewLine & _
          vbNewLine & _
          "'" & strAddWord1 & "'    " & vbNewLine & _
          vbNewLine & _
          "上記のファイルを移動･削除した場合は、元に戻して下さい    " & vbNewLine & _
          "また、ファイル名が異なる場合は、修正して下さい"
  Case 31
    msg = msg & _
          "抽出元または抽出先のシートが見当たりません    " & vbNewLine & _
          "シート名を変更した場合は、元に戻して下さい"
  Case 32
    msg = msg & _
          "同じメーカーの硝種が2つ以上存在します" & vbNewLine & _
          "硝種は各メーカー毎に1つのみを指定して下さい    "
  Case 99
    msg = msg & _
          "'公差ウィンドウ'ボタンをクリックしてモードを解除して下さい    " & vbNewLine & _
          "解除するまでファイルを閉じたり、保存することは出来ません"
  Case 100
    msg = "一括セーブが正常に終了しました" & vbNewLine & vbNewLine & _
          strAddWord1 & vbNewLine & vbNewLine & _
          strAddWord2 & vbNewLine & vbNewLine & _
          " NOTE : 図面情報の図面名と図面番号が空欄の図面データは、    " & vbNewLine & _
          "           この機能では処理できない為、保存されません"
  Case 101
    msg = "一括リネームが正常に終了しました" & vbNewLine & vbNewLine & _
          strAddWord1 & vbNewLine & vbNewLine & _
          strAddWord2 & vbNewLine & vbNewLine & _
          " NOTE : 図面データサマリと対応が取れない名前のファイルや    " & vbNewLine & _
          "           対応するシートの品目コード欄が未記入のものは、" & vbNewLine & _
          "           リネームしておりません"
  Case 102
    msg = "ファイル名を元に戻すには、一括リネームを実行した際に    " & vbNewLine & _
          "対象フォルダ内へ自動保存されたログファイルが必要です" & vbNewLine & _
          "(デフォルト名：'一括リネーム.log')" & vbNewLine & _
          vbNewLine & _
          "次に表示されるファイル選択のダイアログボックスでは、" & vbNewLine & _
          "このログファイルを指定して下さい" & vbNewLine & _
          vbNewLine & _
          " NOTE : ログファイルが無いと、元に戻せません"
  Case 103
    msg = "一括リネームで変更したファイル名を元に戻しました" & vbNewLine & vbNewLine & _
          strAddWord1 & vbNewLine & vbNewLine & _
          strAddWord2 & vbNewLine & vbNewLine & _
          " NOTE : 手作業で名前を変更した場合等、ログファイルの" & vbNewLine & _
          "           内容と一致しないものは処理をスキップしています    "
  Case 104
    msg = "図面データサマリの変換が完了しました" & vbNewLine & _
          vbNewLine & _
          " NOTE : 既に開いていた指定ファイルを修正しました" & vbNewLine & _
          "           ファイル保存はしていないので、ご注意下さい    "
  Case 200
    msg = "光学有効径が保存されていない面がありました" & vbNewLine & _
          "図面情報へ展開する際は、全硝子に光学有効径が必要です" & vbNewLine & _
          vbNewLine & _
          "'" & DefaultSheet1 & "'シートに手入力する場合は[OK]" & vbNewLine & _
          "作業を中止する場合は[キャンセル]をクリックして下さい" & vbNewLine & _
          vbNewLine & _
          " NOTE : 'CVSEQforDDTOOL.seq'マクロでファイル保存すると    " & vbNewLine & _
          "           光学有効径も自動的に保存されます"
  Case 201
    msg = "品目コードの読込には、次のいずれかのテキスト形式で" & vbNewLine & _
          "所定の一覧表を保存した読込用ファイルが必要です" & vbNewLine & _
          vbNewLine & _
          "    ･ メモ帳等にｺﾋﾟｰ&ﾍﾟｰｽﾄしたタブ区切りテキスト(*.txt)    " & vbNewLine & _
          "    ･ 所定の機能から出力したCSVファイル(*.csv)" & vbNewLine & _
          vbNewLine & _
          "読込用ファイルの準備が整ったら[OK]" & vbNewLine & _
          "作業を中止する場合は[キャンセル]をクリックして下さい" & vbNewLine & _
          vbNewLine & _
          " NOTE : 読込用ファイルの詳細は、取扱説明書をご覧下さい"
  Case 202
    msg = "Config\" & strAddWord1 & ".cfg に上書き保存します。宜しいですか?    "
  Case 203
    msg = "図面データサマリの変換が完了しました" & vbNewLine & _
          "修正内容を元のファイルに上書き保存しますか?" & vbNewLine & _
          vbNewLine & _
          "上書き保存する場合は[OK]" & vbNewLine & _
          "旧版を残したい場合は[キャンセル]をクリックして下さい" & vbNewLine & _
          vbNewLine & _
          " NOTE : キャンセルした場合、元のファイルと同じフォルダに    " & vbNewLine & _
          "               '新版_" & strAddWord1 & "'    " & vbNewLine & _
          "           というファイル名で新規保存されます"
End Select

intRtn = MsgBox(Prompt:=msg, Buttons:=btn, Title:=ProjectTitle)

End Sub


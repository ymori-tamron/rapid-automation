Attribute VB_Name = "basUDF"


Option Explicit
Option Base 1
Option Compare Binary


Public Function Sag(ByVal Radius As Variant, ByVal Diameter As Double, _
                    Optional ByVal Conic As Double = 0, _
                    Optional ByVal A4 As Double = 0, _
                    Optional ByVal A6 As Double = 0, _
                    Optional ByVal A8 As Double = 0, _
                    Optional ByVal A10 As Double = 0, _
                    Optional ByVal A12 As Double = 0, _
                    Optional ByVal A14 As Double = 0) As Variant
Attribute Sag.VB_Description = "サグ量を計算します。"
Attribute Sag.VB_ProcData.VB_Invoke_Func = " \n21"

Dim c As Double, h As Double, arg As Double

If Not IsNumeric(Radius) Or _
   Radius = 0 Then
  c = 0
Else
  c = 1 / Radius
End If

h = Diameter / 2
arg = 1 - (1 + Conic) * (c ^ 2) * (h ^ 2)
If arg >= 0 Then
  Sag = (c * (h ^ 2)) / (1 + Sqr(arg)) + _
        A4 * (h ^ 4) + A6 * (h ^ 6) + A8 * (h ^ 8) + A10 * (h ^ 10) + _
        A12 * (h ^ 12) + A14 * (h ^ 14)
Else
  Sag = "#VALUE!"
End If

End Function


Public Function FocalLength(ByVal Radius1 As Variant, ByVal Radius2 As Variant, _
                            ByVal Thickness As Double, ByVal Index As Variant) As Variant
Attribute FocalLength.VB_Description = "単レンズの焦点距離(無限共役)を計算します。"
Attribute FocalLength.VB_ProcData.VB_Invoke_Func = " \n21"

Dim r1 As Double, r2 As Double, n As Double, pw As Double

If Not IsNumeric(Radius1) Or _
   Radius1 = 0 Then
  r1 = 10000000000#
Else
  r1 = Radius1
End If

If Not IsNumeric(Radius2) Or _
   Radius2 = 0 Then
  r2 = 10000000000#
Else
  r2 = Radius2
End If

If Not IsNumeric(Index) Or _
   Index < 1 Then
  n = 1
Else
  n = Index
End If

pw = (n - 1) * (1 / r1 - 1 / r2 + (Thickness * (n - 1)) / (n * r1 * r2))
If pw = 0 Then
  FocalLength = "Inf"
Else
  FocalLength = 1 / pw
End If

End Function


Public Function GlassWeight(ByVal Radius1 As Variant, ByVal Radius2 As Variant, _
                            ByVal Thickness As Double, ByVal SpG As Double, _
                            ByVal Diameter1 As Double, ByVal Diameter2 As Double, _
                            ByVal MaxDiameter As Double, _
                            Optional ByVal Conic1 As Double = 0, _
                            Optional ByVal A4R1 As Double = 0, _
                            Optional ByVal A6R1 As Double = 0, _
                            Optional ByVal A8R1 As Double = 0, _
                            Optional ByVal A10R1 As Double = 0, _
                            Optional ByVal A12R1 As Double = 0, _
                            Optional ByVal A14R1 As Double = 0, _
                            Optional ByVal Conic2 As Double = 0, _
                            Optional ByVal A4R2 As Double = 0, _
                            Optional ByVal A6R2 As Double = 0, _
                            Optional ByVal A8R2 As Double = 0, _
                            Optional ByVal A10R2 As Double = 0, _
                            Optional ByVal A12R2 As Double = 0, _
                            Optional ByVal A14R2 As Double = 0) As Variant
Attribute GlassWeight.VB_Description = "単レンズの重量を計算します。"
Attribute GlassWeight.VB_ProcData.VB_Invoke_Func = " \n21"

Dim chkErr As Variant, r(2) As Variant
Dim i As Long, j As Long, iMax As Long
Dim Pi As Double, dh(2) As Double, v(2) As Double
Dim c(7, 2) As Double, z(0 To 100, 2) As Double

Pi = 4 * Atn(1): iMax = UBound(z, 1)

dh(1) = (Diameter1 / 2) / iMax
dh(2) = (Diameter2 / 2) / iMax

r(1) = Radius1: c(1, 1) = Conic1
c(2, 1) = A4R1: c(3, 1) = A6R1: c(4, 1) = A8R1: c(5, 1) = A10R1: c(6, 1) = A12R1: c(7, 1) = A14R1
r(2) = Radius2: c(1, 2) = Conic2
c(2, 2) = A4R2: c(3, 2) = A6R2: c(4, 2) = A8R2: c(5, 2) = A10R2: c(6, 2) = A12R2: c(7, 2) = A14R2

For j = 1 To 2
  For i = 0 To iMax
    chkErr = (j * (-2) + 3) * _
      Sag(r(j), (dh(j) * i) * 2, c(1, j), c(2, j), c(3, j), c(4, j), c(5, j), c(6, j), c(7, j))
    If Not IsNumeric(chkErr) Then
      GlassWeight = "#VALUE!": Exit Function
    Else
      z(i, j) = chkErr
    End If
    
    If i = 0 Then
      v(j) = 0
    Else
      v(j) = v(j) - (((dh(j) * iMax) ^ 2) * 2 - ((dh(j) * (i - 1)) ^ 2) - _
             ((dh(j) * i) ^ 2)) * Pi * (z(i, j) - z(i - 1, j)) / 2
    End If
  Next
  v(j) = v(j) - (((MaxDiameter / 2) ^ 2) - _
         ((dh(j) * iMax) ^ 2)) * Pi * z(iMax, j)
Next

GlassWeight = _
  SpG * (Pi * ((MaxDiameter / 2) ^ 2) * Thickness + v(1) + v(2)) / 1000

End Function


Public Function NewtonToDeltaRadius(ByVal Newton As Double, ByVal CXorCC As String, _
                                    ByVal Radius As Variant, ByVal Diameter As Double) As Variant
Attribute NewtonToDeltaRadius.VB_Description = "ニュートン公差量を曲率半径の差分量に換算します。"
Attribute NewtonToDeltaRadius.VB_ProcData.VB_Invoke_Func = " \n21"

Dim r As Double, h As Double, arg1 As Double, arg2 As Double

If Not IsNumeric(Radius) Or _
   Radius = 0 Then
  NewtonToDeltaRadius = "Inf": Exit Function
Else
  r = Abs(Radius)
End If

h = Diameter / 2
If (r ^ 2) - (h ^ 2) >= 0 Then
  arg1 = r - Sqr((r ^ 2) - (h ^ 2))
  arg2 = (632.8 / 1000000) * (Newton / 2)
Else
  NewtonToDeltaRadius = "#VALUE!": Exit Function
End If

Select Case CXorCC
  Case "CC"
    If arg1 - arg2 <> 0 Then
      NewtonToDeltaRadius = ((-r) + (1 / 2) * ((h ^ 2) / (arg1 - arg2) + (arg1 - arg2))) * 1000
    Else
      NewtonToDeltaRadius = "#VALUE!": Exit Function
    End If
  Case "CX"
    If arg1 + arg2 <> 0 Then
      NewtonToDeltaRadius = ((r) - (1 / 2) * ((h ^ 2) / (arg1 + arg2) + (arg1 + arg2))) * 1000
    Else
      NewtonToDeltaRadius = "#VALUE!": Exit Function
    End If
  Case Else
    NewtonToDeltaRadius = "#VALUE!": Exit Function
End Select

End Function


Public Function PressValue(ByVal Radius1 As Variant, ByVal Radius2 As Variant, _
                           ByVal Diameter1 As Double, ByVal Diameter2 As Double, _
                           ByVal MaxDiameter As Double) As Double
Attribute PressValue.VB_Description = "単レンズの芯取り代(プレス径－レンズ径)を直径の差分で出力します。"
Attribute PressValue.VB_ProcData.VB_Invoke_Func = " \n21"

Dim objPressParam As clsPressParam

Dim i As Long, zz As Long
Dim r1 As Double, r2 As Double
Dim zVal(7) As Double

If Not IsNumeric(Radius1) Or _
   Radius1 = 0 Then
  r1 = 10000000000#
Else
  r1 = Radius1
End If

If Not IsNumeric(Radius2) Or _
   Radius2 = 0 Then
  r2 = 10000000000#
Else
  r2 = Radius2
End If

zVal(1) = Abs(Diameter1 / r1 - Diameter2 / r2) / 4
zVal(2) = 0.04
zVal(3) = 0.06
zVal(4) = 0.1
zVal(5) = 0.15
zVal(6) = 0.2
zVal(7) = 10000000000#
For i = 2 To UBound(zVal)
  If zVal(1) < zVal(i) Then
    zz = i - 1: Exit For
  End If
Next

Set objPressParam = New clsPressParam
With objPressParam
  .Import strConfig:="PressParam"
  For i = 1 To .lngMaxCount("DP寸法")
    If MaxDiameter < .dbldDP(i, 9) Then
      PressValue = .dbldDP(i, zz)
      Exit For
    End If
  Next
End With
Set objPressParam = Nothing

End Function


Public Sub Load_UDFHelp()

Dim argStr_Sag(9) As String
Dim argStr_FocalLength(4) As String
Dim argStr_GlassWeight(21) As String
Dim argStr_NewtonToDeltaRadius(4) As String
Dim argStr_PressValue(5) As String

argStr_Sag(1) = "には、曲率半径を指定します。(必須)"
argStr_Sag(2) = "には、サグ量を計算する高さを直径で指定します。(必須)"
argStr_Sag(3) = "には、コーニック定数を指定します。(省略可)"
argStr_Sag(4) = "には、4次の非球面係数を指定します。(省略可)"
argStr_Sag(5) = "には、6次の非球面係数を指定します。(省略可)"
argStr_Sag(6) = "には、8次の非球面係数を指定します。(省略可)"
argStr_Sag(7) = "には、10次の非球面係数を指定します。(省略可)"
argStr_Sag(8) = "には、12次の非球面係数を指定します。(省略可)"
argStr_Sag(9) = "には、14次の非球面係数を指定します。(省略可)"

argStr_FocalLength(1) = "には、R1面の曲率半径を指定します。(必須)"
argStr_FocalLength(2) = "には、R2面の曲率半径を指定します。(必須)"
argStr_FocalLength(3) = "には、中心面間隔を指定します。(必須)"
argStr_FocalLength(4) = "には、硝材の屈折率を指定します。(必須)"

argStr_GlassWeight(1) = "には、R1面の曲率半径を指定します。(必須)"
argStr_GlassWeight(2) = "には、R2面の曲率半径を指定します。(必須)"
argStr_GlassWeight(3) = "には、中心面間隔を指定します。(必須)"
argStr_GlassWeight(4) = "には、硝材の比重を指定します。(必須)"
argStr_GlassWeight(5) = "には、R1面の研磨面径を直径で指定します。(必須)"
argStr_GlassWeight(6) = "には、R2面の研磨面径を直径で指定します。(必須)"
argStr_GlassWeight(7) = "には、最大外径を直径で指定します。(必須)"
argStr_GlassWeight(8) = "には、R1面のコーニック定数を指定します。(省略可)"
argStr_GlassWeight(9) = "には、R1面の4次の非球面係数を指定します。(省略可)"
argStr_GlassWeight(10) = "には、R1面の6次の非球面係数を指定します。(省略可)"
argStr_GlassWeight(11) = "には、R1面の8次の非球面係数を指定します。(省略可)"
argStr_GlassWeight(12) = "には、R1面の10次の非球面係数を指定します。(省略可)"
argStr_GlassWeight(13) = "には、R1面の12次の非球面係数を指定します。(省略可)"
argStr_GlassWeight(14) = "には、R1面の14次の非球面係数を指定します。(省略可)"
argStr_GlassWeight(15) = "には、R2面のコーニック定数を指定します。(省略可)"
argStr_GlassWeight(16) = "には、R2面の4次の非球面係数を指定します。(省略可)"
argStr_GlassWeight(17) = "には、R2面の6次の非球面係数を指定します。(省略可)"
argStr_GlassWeight(18) = "には、R2面の8次の非球面係数を指定します。(省略可)"
argStr_GlassWeight(19) = "には、R2面の10次の非球面係数を指定します。(省略可)"
argStr_GlassWeight(20) = "には、R2面の12次の非球面係数を指定します。(省略可)"
argStr_GlassWeight(21) = "には、R2面の14次の非球面係数を指定します。(省略可)"

argStr_NewtonToDeltaRadius(1) = "には、ニュートン本数を指定します。(必須)"
argStr_NewtonToDeltaRadius(2) = "には、面の凸凹を識別するキーワード(凸=>" & _
                                Chr(34) & "CX" & Chr(34) & ", 凹=>" & _
                                Chr(34) & "CC" & Chr(34) & ")を指定します。(必須)"
argStr_NewtonToDeltaRadius(3) = "には、曲率半径の設計値を指定します。(必須)"
argStr_NewtonToDeltaRadius(4) = "には、研磨面径の設計値を直径で指定します。(必須)"

argStr_PressValue(1) = "には、R1面の曲率半径を指定します。(必須)"
argStr_PressValue(2) = "には、R2面の曲率半径を指定します。(必須)"
argStr_PressValue(3) = "には、R1面の研磨面径を直径で指定します。(必須)"
argStr_PressValue(4) = "には、R2面の研磨面径を直径で指定します。(必須)"
argStr_PressValue(5) = "には、最大外径を直径で指定します。(必須)"

Application.ScreenUpdating = False
ThisWorkbook.IsAddin = False

#If VBA7 Then
  Application.MacroOptions _
    Macro:="Sag", _
    Description:="サグ量を計算します。", _
    Category:=ProjectTitle, _
    ArgumentDescriptions:=argStr_Sag
  
  Application.MacroOptions _
    Macro:="FocalLength", _
    Description:="単レンズの焦点距離(無限共役)を計算します。", _
    Category:=ProjectTitle, _
    ArgumentDescriptions:=argStr_FocalLength
  
  Application.MacroOptions _
    Macro:="GlassWeight", _
    Description:="単レンズの重量を計算します。", _
    Category:=ProjectTitle, _
    ArgumentDescriptions:=argStr_GlassWeight
  
  Application.MacroOptions _
    Macro:="NewtonToDeltaRadius", _
    Description:="ニュートン公差量を曲率半径の差分量に換算します。", _
    Category:=ProjectTitle, _
    ArgumentDescriptions:=argStr_NewtonToDeltaRadius
  
  Application.MacroOptions _
    Macro:="PressValue", _
    Description:="単レンズの芯取り代(プレス径－レンズ径)を直径の差分で出力します。", _
    Category:=ProjectTitle, _
    ArgumentDescriptions:=argStr_PressValue
#Else
  Application.MacroOptions _
    Macro:="Sag", _
    Description:="サグ量を計算します。", _
    Category:=ProjectTitle
  
  Application.MacroOptions _
    Macro:="FocalLength", _
    Description:="単レンズの焦点距離(無限共役)を計算します。", _
    Category:=ProjectTitle
  
  Application.MacroOptions _
    Macro:="GlassWeight", _
    Description:="単レンズの重量を計算します。", _
    Category:=ProjectTitle
  
  Application.MacroOptions _
    Macro:="NewtonToDeltaRadius", _
    Description:="ニュートン公差量を曲率半径の差分量に換算します。", _
    Category:=ProjectTitle
  
  Application.MacroOptions _
    Macro:="PressValue", _
    Description:="単レンズの芯取り代(プレス径－レンズ径)を直径の差分で出力します。", _
    Category:=ProjectTitle
#End If

ThisWorkbook.IsAddin = True
ThisWorkbook.Saved = True
Application.ScreenUpdating = True

End Sub


VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsProgCtrl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False


Option Explicit
Option Base 1
Option Compare Binary

Private lngMC(2) As Long
Private dblMW As Double, dblCW(2) As Double


Private Sub Class_Initialize()

frmProg.Show
dblMW = 153
dblCW(1) = 0: dblCW(2) = 0
lngMC(1) = 99: lngMC(2) = 99

End Sub


Private Sub Class_Terminate()

apiSleep dwMSec:=500
Unload Object:=frmProg

End Sub


Public Sub InitBar1(ByVal lngMaxCount As Long)

lngMC(1) = lngMaxCount
dblCW(1) = 0

End Sub


Public Sub InitBar2(ByVal lngMaxCount As Long)

lngMC(2) = lngMaxCount
dblCW(2) = 0

apiSleep dwMSec:=200
With frmProg
  .ProgBar2.Width = 0
  .ProgRatio2.Caption = "0%"
  .Label2.Caption = Left(ActiveSheet.Name, 3)
  .Repaint
End With

End Sub


Public Sub AddBar(ByVal lngDivCount As Long)

With frmProg
  dblCW(1) = dblCW(1) + (dblMW / (lngMC(1) * lngMC(2))) / lngDivCount
  If dblCW(1) > dblMW Then
    .ProgBar1.Width = dblMW
  Else
    .ProgBar1.Width = dblCW(1)
  End If
  .ProgRatio1.Caption = Format(dblCW(1) / dblMW, "0%")
  .ProgFrame1.Repaint
  
  dblCW(2) = dblCW(2) + (dblMW / lngMC(2)) / lngDivCount
  If dblCW(2) > dblMW Then
    .ProgBar2.Width = dblMW
  Else
    .ProgBar2.Width = dblCW(2)
  End If
  .ProgRatio2.Caption = Format(dblCW(2) / dblMW, "0%")
  .ProgFrame2.Repaint
End With

End Sub


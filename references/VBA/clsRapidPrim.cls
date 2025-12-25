VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsRapidPrim"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False


Option Explicit
Option Base 1
Option Compare Binary


Public Sub SetUp(ByRef prmObject As zwDrawCAD.Primitive, ByVal lngLayerNo As Long, _
                 ByVal lngPenStyle As Long, ByVal lngPenWidth As Long)

With rpdDoc.DrawLayers
  If .GetLayer(lNumber:=lngLayerNo) = -1 Then
    .Add lNumber:=lngLayerNo
  End If
End With

With prmObject
  .LayerNo = lngLayerNo
  .PenStyle = lngPenStyle
  .PenWidth = lngPenWidth
End With

End Sub


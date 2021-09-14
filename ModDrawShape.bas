Attribute VB_Name = "ModDrawShape"
Option Explicit

'シェイプ作図関連モジュール
'20210914作成

Function DrawCurve(XYList, TargetSheet As Worksheet) As Shape
'XY座標から曲線を描く
'シェイプをオブジェクト変数として返す
'20210914

'引数
'XYList         ・・・XY座標が入った二次元配列 X方向→右方向 Y方向→下方向
'TargetSheet    ・・・作図対象のシート

    Dim I%, Count%
    Count = UBound(XYList, 1)
    
    With TargetSheet.Shapes.BuildFreeform(msoEditingCorner, XYList(1, 1), XYList(1, 2))
        
        For I = 2 To Count
            .AddNodes msoSegmentCurve, msoEditingAuto, XYList(I, 1), XYList(I, 2)
        Next I
        Set DrawCurve = .ConvertToShape
    End With
    
End Function

Sub AddPointToCurve(InputShape As Shape, AddX#, AddY#, Optional DeleteFirstPoint As Boolean = True)
'指定点まで曲線を延長する
'20210914

'引数
'InputShape         ・・・対象の曲線
'AddX               ・・・追加する点のX座標（右方向）
'AddY               ・・・追加する点のY座標（下方向）
'[DeleteFirstPoint] ・・・対象の曲線の最初の点を削除するかどうか


    Dim TmpNode As ShapeNodes
    Set TmpNode = InputShape.Nodes
    
    With TmpNode
        .Insert .Count, msoSegmentCurve, msoEditingSmooth, AddX, AddY
        TmpNode.SetEditingType .Count - 5, msoEditingAuto
    End With
    
    If DeleteFirstPoint Then
        TmpNode.Delete 1
    End If
    
End Sub




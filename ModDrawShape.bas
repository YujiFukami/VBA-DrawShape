Attribute VB_Name = "ModDrawShape"
Option Explicit

'DrawCurve           ・・・元場所：FukamiAddins3.ModDrawShape
'DrawCurveAddPoint   ・・・元場所：FukamiAddins3.ModDrawShape
'DrawPolyLine        ・・・元場所：FukamiAddins3.ModDrawShape
'DrawPolyLineAddPoint・・・元場所：FukamiAddins3.ModDrawShape



Function DrawCurve(XYList, TargetSheet As Worksheet) As Shape
'XY座標から曲線を描く
'シェイプをオブジェクト変数として返す
'20210914

'引数
'XYList         ・・・XY座標が入った二次元配列 X方向→右方向 Y方向→下方向
'TargetSheet    ・・・作図対象のシート

    Dim I     As Integer
    Dim Count As Integer
    Count = UBound(XYList, 1)
    
    With TargetSheet.Shapes.BuildFreeform(msoEditingCorner, XYList(1, 1), XYList(1, 2))
        
        For I = 2 To Count
            .AddNodes msoSegmentCurve, msoEditingAuto, XYList(I, 1), XYList(I, 2)
        Next I
        Set DrawCurve = .ConvertToShape
    End With
    
End Function

Sub DrawCurveAddPoint(InputShape As Shape, AddX As Double, AddY As Double, Optional DeleteFirstPoint As Boolean = True)
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

Function DrawPolyLine(XYList, TargetSheet As Worksheet) As Shape
'XY座標からポリラインを描く
'シェイプをオブジェクト変数として返す
'20210921

'引数
'XYList         ・・・XY座標が入った二次元配列 X方向→右方向 Y方向→下方向
'TargetSheet    ・・・作図対象のシート

    Dim I     As Integer
    Dim Count As Integer
    Count = UBound(XYList, 1)
    
    With TargetSheet.Shapes.BuildFreeform(msoEditingCorner, XYList(1, 1), XYList(1, 2))
        
        For I = 2 To Count
            .AddNodes msoSegmentLine, msoEditingAuto, XYList(I, 1), XYList(I, 2)
        Next I
        Set DrawPolyLine = .ConvertToShape
    End With
    
End Function

Sub DrawPolyLineAddPoint(InputShape As Shape, AddX As Double, AddY As Double, Optional DeleteFirstPoint As Boolean = True)
'ポリラインに点を追加して延長する
'20211008

'引数
'InputShape         ・・・対象のポリライン
'AddX               ・・・追加する点のX座標（右方向）
'AddY               ・・・追加する点のY座標（下方向）
'[DeleteFirstPoint] ・・・対象の曲線の最初の点を削除するかどうか

    Dim TmpNode As ShapeNodes
    Set TmpNode = InputShape.Nodes
    
    With TmpNode
        .Insert .Count, msoSegmentLine, msoEditingCorner, AddX, AddY
    End With
    
    If DeleteFirstPoint Then
        TmpNode.Delete 1
    End If
    
End Sub



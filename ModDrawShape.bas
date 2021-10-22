Attribute VB_Name = "ModDrawShape"
Option Explicit

'DrawCurve           �E�E�E���ꏊ�FFukamiAddins3.ModDrawShape
'DrawCurveAddPoint   �E�E�E���ꏊ�FFukamiAddins3.ModDrawShape
'DrawPolyLine        �E�E�E���ꏊ�FFukamiAddins3.ModDrawShape
'DrawPolyLineAddPoint�E�E�E���ꏊ�FFukamiAddins3.ModDrawShape



Function DrawCurve(XYList, TargetSheet As Worksheet) As Shape
'XY���W����Ȑ���`��
'�V�F�C�v���I�u�W�F�N�g�ϐ��Ƃ��ĕԂ�
'20210914

'����
'XYList         �E�E�EXY���W���������񎟌��z�� X�������E���� Y������������
'TargetSheet    �E�E�E��}�Ώۂ̃V�[�g

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
'�w��_�܂ŋȐ�����������
'20210914

'����
'InputShape         �E�E�E�Ώۂ̋Ȑ�
'AddX               �E�E�E�ǉ�����_��X���W�i�E�����j
'AddY               �E�E�E�ǉ�����_��Y���W�i�������j
'[DeleteFirstPoint] �E�E�E�Ώۂ̋Ȑ��̍ŏ��̓_���폜���邩�ǂ���


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
'XY���W����|�����C����`��
'�V�F�C�v���I�u�W�F�N�g�ϐ��Ƃ��ĕԂ�
'20210921

'����
'XYList         �E�E�EXY���W���������񎟌��z�� X�������E���� Y������������
'TargetSheet    �E�E�E��}�Ώۂ̃V�[�g

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
'�|�����C���ɓ_��ǉ����ĉ�������
'20211008

'����
'InputShape         �E�E�E�Ώۂ̃|�����C��
'AddX               �E�E�E�ǉ�����_��X���W�i�E�����j
'AddY               �E�E�E�ǉ�����_��Y���W�i�������j
'[DeleteFirstPoint] �E�E�E�Ώۂ̋Ȑ��̍ŏ��̓_���폜���邩�ǂ���

    Dim TmpNode As ShapeNodes
    Set TmpNode = InputShape.Nodes
    
    With TmpNode
        .Insert .Count, msoSegmentLine, msoEditingCorner, AddX, AddY
    End With
    
    If DeleteFirstPoint Then
        TmpNode.Delete 1
    End If
    
End Sub



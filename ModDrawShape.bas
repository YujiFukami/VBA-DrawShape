Attribute VB_Name = "ModDrawShape"
Option Explicit

'�V�F�C�v��}�֘A���W���[��
'20210914�쐬

Function DrawCurve(XYList, TargetSheet As Worksheet) As Shape
'XY���W����Ȑ���`��
'�V�F�C�v���I�u�W�F�N�g�ϐ��Ƃ��ĕԂ�
'20210914

'����
'XYList         �E�E�EXY���W���������񎟌��z�� X�������E���� Y������������
'TargetSheet    �E�E�E��}�Ώۂ̃V�[�g

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




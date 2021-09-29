Attribute VB_Name = "ModCell"
Option Explicit


Sub SelectA1()
'�S�V�[�g��A1�Z����I������
'20210720

    Dim TmpSheet As Worksheet
    For Each TmpSheet In ActiveWorkbook.Sheets
        Application.GoTo TmpSheet.Range("A1")
    Next
    
End Sub

Function GetBlankCell(TargetCell As Range)
'�w��V�[�g���̋󔒃Z�����擾����
'�֐��v���o���p
'20210720
    
    Dim Output As Range
    Set Output = TargetCell.SpecialCells(xlCellTypeBlanks)
    Set GetBlankCell = Output
    
End Function

Sub SortCell(TargetCell As Range, KeyCell As Range, Optional InputOrder As OrderType = xlAscending)
'�w��͈͂̃Z������ёւ���
'20210720

'TargetCell�E�E�E���ёւ��͈͂̃Z��
'KeyCell   �E�E�E���ёւ��̃L�[�ƂȂ�Z��
'InputOrder�E�E�E����(xlAscending)���~��(xlDescending) �f�t�H���g�Ȃ珸��

    Dim TargetSheet As Worksheet
    Set TargetSheet = TargetCell.Parent
    
    With TargetSheet.Sort.SortFields
        .Clear
        .Add Key:=KeyCell, _
             SortOn:=xlSortOnValues, _
             Order:=InputOrder, _
             DataOption:=xlSortNormal
    End With
    
    With TargetSheet.Sort
        .SetRange TargetCell
        .Header = xlNo '�擪�s�̓w�b�_�[�łȂ�
        .Orientation = xlTopToBottom
        .Apply
    End With
End Sub

Sub SetCommentPicture(TargetCell As Range, PicturePath$)
'�Z���̃R�����g�ŉ摜��\������
'20210720

    If Dir(PicturePath, vbDirectory) = "" Then
        MsgBox ("�摜�t�@�C���u " & PicturePath & "�v��������܂���" & vbLf & _
               "�������L�����Z�����܂�")
        Exit Sub
    End If
    
    Dim Img As Object
    Set Img = LoadPicture(PicturePath)
    
    With TargetCell.AddComment
        .Shape.Fill.UserPicture PicturePath
        .Shape.Height = Application.CentimetersToPoints(Img.Height) / 1000
        .Shape.Width = Application.CentimetersToPoints(Img.Width) / 1000
        .Visible = True
    End With
    
End Sub

Sub ResetFilter(InputSheet As Worksheet)
'�w��V�[�g�̃t�B���^����������B
'InputSheet0�E�E�E���s�ΏۃV�[�g�B�u�I�u�W�F�N�g�v�A�u�V�[�g���v�A�u�V�[�g�ԍ��v�ǂ�Ŏw�肵�Ă��悢�B�����͂Ȃ猻�݃A�N�e�B�u�ȃV�[�g�B
'�Q�l��http://officetanaka.net/excel/vba/tips/tips129.htm
'20210721

    Dim I&
    
    InputSheet.Select
    If ActiveSheet.AutoFilterMode Then '�I�[�g�t�B���^���ݒ肳��Ă���ꍇ
        For I = 1 To InputSheet.AutoFilter.Filters.Count '���̗�𒲍�����
            If InputSheet.AutoFilter.Filters(I).On Then '�t�B���^���ݒ肳��Ă���ꍇ
                InputSheet.Select
                Selection.AutoFilter Field:=I '�t�B���^���ݒ肳��Ă����̃t�B���^����
                
            End If
        Next I
    End If

End Sub

Function GetEndRow&(StartCell As Range, Optional ByVal MaxRenzokuBlank& = 0)
'�I�[�g�t�B���^���ݒ肵�Ă���ꍇ���l�����Ă̍ŏI�s�̎擾
'20210728

'StartCell          :�T�������̊J�n�Z��
'MaxRenzokuBlank    :�󔒃Z���̘A����(�����ȏ�̋󔒃Z�����A��������A�Ō�̔�󔒃Z�����ŏI�Z��)
    
    Dim InputSheet As Worksheet, OutputSheet As Worksheet '���o�̓V�[�g
    Set InputSheet = StartCell.Parent
    
    Dim StartRow&, StartCol&
    Dim TmpRenzokuBlank&, TmpEndRow&
    Dim TmpRow&
    Dim I&, J&, K&, M&, N& '�����グ�p(Long�^)
    If InputSheet.AutoFilterMode Or MaxRenzokuBlank <> 0 Then '�I�[�g�t�B���^���ݒ肳��Ă���ꍇ
        If MaxRenzokuBlank = 0 Then
            MaxRenzokuBlank = 500 '����������������������������������������������
        End If
        StartRow = StartCell.Row
        StartCol = StartCell.Column
        For TmpRow = StartRow To Rows.Count
            If InputSheet.Cells(TmpRow, StartCol).Value = "" Then
                If MaxRenzokuBlank = 0 Then
                    '���̈ʒu�̎�O���ŏI�s
                    Exit For
                Else
                    TmpRenzokuBlank = TmpRenzokuBlank + 1
                End If
                
                If TmpRenzokuBlank > MaxRenzokuBlank Then
                    '�w�肵�����ȏ�ɋ󔒃Z�����A�������ꍇ�́A�Ō�̔�󔒃Z�����ŏI�s
                    Exit For
                End If
            Else
                TmpEndRow = TmpRow
                TmpRenzokuBlank = 0
            End If
        Next TmpRow
    
    Else '�I�[�g�t�B���^���ݒ肳��Ă��Ȃ��ꍇ
        '�ʏ�̍ŏI�s�̎擾
        TmpEndRow = InputSheet.Cells(Rows.Count, StartCell.Column).End(xlUp).Row
    End If
    
    GetEndRow = TmpEndRow '�o��

End Function

Function GetEndCell(StartCell As Range, Optional MaxRenzokuBlank& = 0) As Range
'�I�[�g�t�B���^���ݒ肵�Ă���ꍇ���l�����Ă̍ŏI�Z���̎擾
'20210728

'StartCell          :�T�������̊J�n�Z��
'MaxRenzokuBlank    :�󔒃Z���̘A����(�����ȏ�̋󔒃Z�����A��������A�Ō�̔�󔒃Z�����ŏI�Z��)

    Dim EndRow&
    EndRow = GetEndRow(StartCell, MaxRenzokuBlank)
    Dim InputSheet As Worksheet, OutputSheet As Worksheet '���o�̓V�[�g
    Set InputSheet = StartCell.Parent
    
    Dim Output As Range
    Set Output = InputSheet.Cells(EndRow, StartCell.Column)
    Set GetEndCell = Output
    
End Function

Sub SetCellDataBar(TargetCell As Range, Ratio#, Color&)
'�Z���̏����ݒ��0�`1�̒l�Ɋ�Â��āA�f�[�^�o�[��ݒ肷��
'20210820

'TargetCell :�Ώۂ̃Z��
'Ratio      :�����i0�`1�j
'Color      :�o�[�̐F�iRGB�l�j

    Dim Gosa#
    Gosa = 10 ^ (-10) '����������������������������������������������
    
    With TargetCell
        
        .Interior.Pattern = xlPatternLinearGradient
        .Interior.Gradient.Degree = 0
        
        With .Interior.Gradient.ColorStops
            
            If Ratio > Gosa Then
                .Add(0).Color = Color
                .Add(Gosa).Color = Color
                .Add(Gosa * 2).Color = Color
                
                If Gosa * 3 < Ratio Then
                    .Add(Ratio).Color = Color
                Else
                    .Add(Gosa * 3).Color = Color
                End If
            End If
            
            If Ratio < 1 Then
                If Ratio + Gosa > 1 Then
                    .Add((1 + Ratio) / 2).Color = Color
                Else
                    .Add(Ratio + Gosa).Color = rgbWhite
                End If
                .Add(1).Color = rgbWhite
            End If
        End With
    End With

End Sub

Sub Test_ShowColumns()

    Dim TargetSheet As Worksheet
    Dim ColumnABCList1D
    ColumnABCList1D = Array("C", "E", "Z")
    ColumnABCList1D = Application.Transpose(Application.Transpose(ColumnABCList1D))
    Set TargetSheet = ActiveSheet
    
    Call ShowColumns(ColumnABCList1D, TargetSheet, "Z", True)

End Sub

Sub ShowColumns(ColumnABCList1D, TargetSheet As Worksheet, Optional ByVal MaxColABC$, Optional InputShow As Boolean = True)
'�w���̂ݕ\���ɂ���
'20210917

'����
'ColumnABCList�E�E�E��\���Ώۂ̗񖼂�1�����z�� ��) ("A","B","C")
'TargetSheet  �E�E�E�Ώۂ̃V�[�g
'MaxColABC    �E�E�E��\���֑ؑΏۂ̗�͈͂̍ő��
'InputShow    �E�E�E�w�ߗ��\���Ȃ�True,��\���Ȃ�False�B�f�t�H���g��True
                                                                 
    '�����`�F�b�N
    Call CheckArray1D(ColumnABCList1D, "ColumnABCList1D")
    Call CheckArray1DStart1(ColumnABCList1D, "ColumnABCList1D")
    
    If MaxColABC = "" Then '��\���֑ؑΏۂ̗�͈͂̍ő�񂪎w�肳��Ă��Ȃ��ꍇ�̓V�[�g�̍ŏI��
        MaxColABC = Split(Cells(1, Columns.Count).Address(True, False), "$")(0) '�ŏI��ԍ��̃A���t�@�x�b�g�擾
    End If
    
    Dim I&, J&, K&, M&, N&      '�����グ�p(Long�^)
    Dim ColumnName$             '�\���Ώۂ̗񖼂��܂Ƃ߂�����
    N = UBound(ColumnABCList1D) '�Ώۂ̗�̌�
    ColumnName = ""             '�񖼂܂Ƃ߂̏�����
    For I = 1 To N
        ColumnName = ColumnName & ColumnABCList1D(I) & ":" & ColumnABCList1D(I)
        If I < N Then '�񖼂̍Ōゾ��","�����Ȃ�
            ColumnName = ColumnName & ","
        End If
    Next I
    
    Dim TargetCell As Range                        '�Ώ۔͈͂̃Z���I�u�W�F�N�g
    Set TargetCell = TargetSheet.Range(ColumnName) '�Ώ۔͈͂��Z���I�u�W�F�N�g�Ŏ擾
                                                                                    
    Application.ScreenUpdating = False             '��ʍX�V���������č�����
    
    If InputShow = True Then                                 '�\���ɐ؂�ւ��邩�A��\���ɐ؂�ւ��邩
        TargetSheet.Columns("A:" & MaxColABC).Hidden = True  '�S�̂��\��
        TargetCell.EntireColumn.Hidden = False               '�w�ߗ�̂ݕ\������
    Else
        TargetSheet.Columns("A:" & MaxColABC).Hidden = False '�S�̂��\��
        TargetCell.EntireColumn.Hidden = True                '�w�ߗ�̂ݕ\������
    End If
    
    ActiveWindow.ScrollColumn = 1     '��ԍ��̗�ɃX�N���[�����ĕ\������
    Application.ScreenUpdating = True '��ʍX�V�����̉���
    
End Sub



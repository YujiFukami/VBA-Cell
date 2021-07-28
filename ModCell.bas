Attribute VB_Name = "ModCell"
Option Explicit

Enum OrderType '�����~���̗񋓌^
    xlAscending = 1
    xlDescending = 2
End Enum

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
    Output = TargetCell.SpecialCells(xlCellTypeBlanks)

End Function

Sub CellSort(TargetCell As Range, KeyCell As Range, Optional InputOrder As OrderType = xlAscending)
'�w��͈͂̃Z������ёւ���
'20210720

'TargetCell�E�E�E���ёւ��͈͂̃Z��
'KeyCell�E�E�E���ёւ��̃L�[�ƂȂ�Z��
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
        
    If ActiveSheet.AutoFilterMode Then '�I�[�g�t�B���^���ݒ肳��Ă���ꍇ
        For I = 1 To InputSheet.AutoFilter.Filters.Count '���̗�𒲍�����
            If InputSheet.AutoFilter.Filters(I).On Then '�t�B���^���ݒ肳��Ă���ꍇ
                Selection.AutoFilter Field:=I '�t�B���^���ݒ肳��Ă����̃t�B���^����
                
            End If
        Next I
    End If

End Sub
Function GetEndRow&(StartCell As Range, Optional MaxRenzokuBlank& = 0)
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
    If InputSheet.AutoFilterMode Then '�I�[�g�t�B���^���ݒ肳��Ă���ꍇ
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

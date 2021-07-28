Attribute VB_Name = "ModCell"
Option Explicit

Enum OrderType '昇順降順の列挙型
    xlAscending = 1
    xlDescending = 2
End Enum

Sub SelectA1()
'全シートのA1セルを選択する
'20210720

    Dim TmpSheet As Worksheet
    For Each TmpSheet In ActiveWorkbook.Sheets
        Application.GoTo TmpSheet.Range("A1")
    Next
    
End Sub
Function GetBlankCell(TargetCell As Range)
'指定シート内の空白セルを取得する
'関数思い出し用
'20210720
    
    Dim Output As Range
    Output = TargetCell.SpecialCells(xlCellTypeBlanks)

End Function

Sub CellSort(TargetCell As Range, KeyCell As Range, Optional InputOrder As OrderType = xlAscending)
'指定範囲のセルを並び替える
'20210720

'TargetCell・・・並び替え範囲のセル
'KeyCell・・・並び替えのキーとなるセル
'InputOrder・・・昇順(xlAscending)か降順(xlDescending) デフォルトなら昇順

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
        .Header = xlNo '先頭行はヘッダーでない
        .Orientation = xlTopToBottom
        .Apply
    End With
End Sub

Sub SetCommentPicture(TargetCell As Range, PicturePath$)
'セルのコメントで画像を表示する
'20210720

    If Dir(PicturePath, vbDirectory) = "" Then
        MsgBox ("画像ファイル「 " & PicturePath & "」が見つかりません" & vbLf & _
               "処理をキャンセルします")
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
'指定シートのフィルタを解除する。
'InputSheet0・・・実行対象シート。「オブジェクト」、「シート名」、「シート番号」どれで指定してもよい。未入力なら現在アクティブなシート。
'参考→http://officetanaka.net/excel/vba/tips/tips129.htm
'20210721

    Dim I&
        
    If ActiveSheet.AutoFilterMode Then 'オートフィルタが設定されている場合
        For I = 1 To InputSheet.AutoFilter.Filters.Count '一つ一つの列を調査して
            If InputSheet.AutoFilter.Filters(I).On Then 'フィルタが設定されている場合
                Selection.AutoFilter Field:=I 'フィルタが設定されている列のフィルタ解除
                
            End If
        Next I
    End If

End Sub
Function GetEndRow&(StartCell As Range, Optional MaxRenzokuBlank& = 0)
'オートフィルタが設定してある場合も考慮しての最終行の取得
'20210728

'StartCell          :探索する基準の開始セル
'MaxRenzokuBlank    :空白セルの連続個数(いくつ以上の空白セルが連続したら、最後の非空白セルが最終セル)
    
    Dim InputSheet As Worksheet, OutputSheet As Worksheet '入出力シート
    Set InputSheet = StartCell.Parent
    
    Dim StartRow&, StartCol&
    Dim TmpRenzokuBlank&, TmpEndRow&
    Dim TmpRow&
    Dim I&, J&, K&, M&, N& '数え上げ用(Long型)
    If InputSheet.AutoFilterMode Then 'オートフィルタが設定されている場合
        StartRow = StartCell.Row
        StartCol = StartCell.Column
        For TmpRow = StartRow To Rows.Count
            If InputSheet.Cells(TmpRow, StartCol).Value = "" Then
                If MaxRenzokuBlank = 0 Then
                    'その位置の手前が最終行
                    Exit For
                Else
                    TmpRenzokuBlank = TmpRenzokuBlank + 1
                End If
                
                If TmpRenzokuBlank > MaxRenzokuBlank Then
                    '指定した数以上に空白セルが連続した場合は、最後の非空白セルが最終行
                    Exit For
                End If
            Else
                TmpEndRow = TmpRow
                TmpRenzokuBlank = 0
            End If
        Next TmpRow
    
    Else 'オートフィルタが設定されていない場合
        '通常の最終行の取得
        TmpEndRow = InputSheet.Cells(Rows.Count, StartCell.Column).End(xlUp).Row
    End If
    
    GetEndRow = TmpEndRow '出力

End Function
Function GetEndCell(StartCell As Range, Optional MaxRenzokuBlank& = 0) As Range
'オートフィルタが設定してある場合も考慮しての最終セルの取得
'20210728

'StartCell          :探索する基準の開始セル
'MaxRenzokuBlank    :空白セルの連続個数(いくつ以上の空白セルが連続したら、最後の非空白セルが最終セル)

    Dim EndRow&
    EndRow = GetEndRow(StartCell, MaxRenzokuBlank)
    Dim InputSheet As Worksheet, OutputSheet As Worksheet '入出力シート
    Set InputSheet = StartCell.Parent
    
    Dim Output As Range
    Set Output = InputSheet.Cells(EndRow, StartCell.Column)
    Set GetEndCell = Output
    
End Function

Attribute VB_Name = "ModCell"
Option Explicit

'SelectA1          ・・・元場所：FukamiAddins3.ModCell 
'GetBlankCell      ・・・元場所：FukamiAddins3.ModCell 
'SortCell          ・・・元場所：FukamiAddins3.ModCell 
'SetCommentPicture ・・・元場所：FukamiAddins3.ModCell 
'ResetFilter       ・・・元場所：FukamiAddins3.ModCell 
'GetEndRow         ・・・元場所：FukamiAddins3.ModCell 
'GetEndCell        ・・・元場所：FukamiAddins3.ModCell 
'SetCellDataBar    ・・・元場所：FukamiAddins3.ModCell 
'Test_ShowColumns  ・・・元場所：FukamiAddins3.ModCell 
'ShowColumns       ・・・元場所：FukamiAddins3.ModCell 
'CheckArray1D      ・・・元場所：FukamiAddins3.ModArray
'CheckArray1DStart1・・・元場所：FukamiAddins3.ModArray

'宣言セクション※※※※※※※※※※※※※※※※※※※※※※※※※※※
'-----------------------------------
'元場所:FukamiAddins3.ModEnum.OrderType
Public Enum OrderType '昇順降順の列挙型
    xlAscending = 1
    xlDescending = 2
End Enum
'宣言セクション終了※※※※※※※※※※※※※※※※※※※※※※※※※※※

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
    Set Output = TargetCell.SpecialCells(xlCellTypeBlanks)
    Set GetBlankCell = Output
    
End Function

Sub SortCell(TargetCell As Range, KeyCell As Range, Optional InputOrder As OrderType = xlAscending)
'指定範囲のセルを並び替える
'20210720

'TargetCell・・・並び替え範囲のセル
'KeyCell   ・・・並び替えのキーとなるセル
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

Sub SetCommentPicture(TargetCell As Range, PicturePath As String)
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

    Dim I As Long
    
    InputSheet.Select
    If ActiveSheet.AutoFilterMode Then 'オートフィルタが設定されている場合
        For I = 1 To InputSheet.AutoFilter.Filters.Count '一つ一つの列を調査して
            If InputSheet.AutoFilter.Filters(I).On Then 'フィルタが設定されている場合
                InputSheet.Select
                Selection.AutoFilter Field:=I 'フィルタが設定されている列のフィルタ解除
                
            End If
        Next I
    End If

End Sub

Function GetEndRow(StartCell As Range, Optional ByVal MaxRenzokuBlank As Long = 0)
'オートフィルタが設定してある場合も考慮しての最終行の取得
'20210728

'StartCell          :探索する基準の開始セル
'MaxRenzokuBlank    :空白セルの連続個数(いくつ以上の空白セルが連続したら、最後の非空白セルが最終セル)
    
    Dim InputSheet      As Worksheet
    Dim StartRow        As Long
    Dim StartCol        As Long
    Dim TmpRenzokuBlank As Long
    Dim TmpEndRow       As Long
    Dim TmpRow          As Long
    Set InputSheet = StartCell.Parent
    If InputSheet.AutoFilterMode Or MaxRenzokuBlank <> 0 Then 'オートフィルタが設定されている場合
        If MaxRenzokuBlank = 0 Then
            MaxRenzokuBlank = 500 '←←←←←←←←←←←←←←←←←←←←←←←
        End If
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

Function GetEndCell(StartCell As Range, Optional MaxRenzokuBlank As Long = 0) As Range
'オートフィルタが設定してある場合も考慮しての最終セルの取得
'20210728

'StartCell          :探索する基準の開始セル
'MaxRenzokuBlank    :空白セルの連続個数(いくつ以上の空白セルが連続したら、最後の非空白セルが最終セル)

    Dim EndRow     As Long
    Dim InputSheet As Worksheet
    EndRow = GetEndRow(StartCell, MaxRenzokuBlank)
    Set InputSheet = StartCell.Parent
    
    Dim Output As Range
    Set Output = InputSheet.Cells(EndRow, StartCell.Column)
    Set GetEndCell = Output
    
End Function

Sub SetCellDataBar(TargetCell As Range, Ratio As Double, Color As Long)
'セルの書式設定で0～1の値に基づいて、データバーを設定する
'20210820

'TargetCell :対象のセル
'Ratio      :割合（0～1）
'Color      :バーの色（RGB値）

    Dim Gosa As Double
    Gosa = 10 ^ (-10) '←←←←←←←←←←←←←←←←←←←←←←←
    
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

    Dim TargetSheet    As Worksheet
    Dim ColumnABCList1D
    ColumnABCList1D = Array("C", "E", "Z")
    ColumnABCList1D = Application.Transpose(Application.Transpose(ColumnABCList1D))
    Set TargetSheet = ActiveSheet
    
    Call ShowColumns(ColumnABCList1D, TargetSheet, "Z", True)

End Sub

Sub ShowColumns(ColumnABCList1D, TargetSheet As Worksheet, Optional ByVal MaxColABC As String, Optional InputShow As Boolean = True)
'指定列のみ表示にする
'20210917

'引数
'ColumnABCList・・・非表示対象の列名の1次元配列 例) ("A","B","C")
'TargetSheet  ・・・対象のシート
'MaxColABC    ・・・非表示切替対象の列範囲の最大列
'InputShow    ・・・指令列を表示ならTrue,非表示ならFalse。デフォルトはTrue
                                                                 
    '引数チェック
    Call CheckArray1D(ColumnABCList1D, "ColumnABCList1D")
    Call CheckArray1DStart1(ColumnABCList1D, "ColumnABCList1D")
    
    If MaxColABC = "" Then '非表示切替対象の列範囲の最大列が指定されていない場合はシートの最終列
        MaxColABC = Split(Cells(1, Columns.Count).Address(True, False), "$")(0) '最終列番号のアルファベット取得
    End If
    
    Dim I          As Long
    Dim N          As Long
    Dim ColumnName As String    '表示対象の列名をまとめたもの
    N = UBound(ColumnABCList1D) '対象の列の個数
    ColumnName = ""             '列名まとめの初期化
    For I = 1 To N
        ColumnName = ColumnName & ColumnABCList1D(I) & ":" & ColumnABCList1D(I)
        If I < N Then '列名の最後だけ","をつけない
            ColumnName = ColumnName & ","
        End If
    Next I
    
    Dim TargetCell As Range                        '対象範囲のセルオブジェクト
    Set TargetCell = TargetSheet.Range(ColumnName) '対象範囲をセルオブジェクトで取得
                                                                                    
    Application.ScreenUpdating = False             '画面更新を解除して高速化
    
    If InputShow = True Then                                 '表示に切り替えるか、非表示に切り替えるか
        TargetSheet.Columns("A:" & MaxColABC).Hidden = True  '全体を非表示
        TargetCell.EntireColumn.Hidden = False               '指令列のみ表示する
    Else
        TargetSheet.Columns("A:" & MaxColABC).Hidden = False '全体を非表示
        TargetCell.EntireColumn.Hidden = True                '指令列のみ表示する
    End If
    
    ActiveWindow.ScrollColumn = 1     '一番左の列にスクロールして表示する
    Application.ScreenUpdating = True '画面更新解除の解除
    
End Sub

Private Sub CheckArray1D(InputArray, Optional HairetuName As String = "配列")
'入力配列が1次元配列かどうかチェックする
'20210804

    Dim Dummy As Integer
    On Error Resume Next
    Dummy = UBound(InputArray, 2)
    On Error GoTo 0
    If Dummy <> 0 Then
        MsgBox (HairetuName & "は1次元配列を入力してください")
        Stop
        Exit Sub '入力元のプロシージャを確認するために抜ける
    End If

End Sub

Private Sub CheckArray1DStart1(InputArray, Optional HairetuName As String = "配列")
'入力1次元配列の開始番号が1かどうかチェックする
'20210804

    If LBound(InputArray, 1) <> 1 Then
        MsgBox (HairetuName & "の開始要素番号は1にしてください")
        Stop
        Exit Sub '入力元のプロシージャを確認するために抜ける
    End If

End Sub



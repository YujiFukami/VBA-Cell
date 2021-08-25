# VBA-Cell
# セル操作関係のVBA
- License: The MIT license

- Copyright (c) 2021 YujiFukami

- 開発テスト環境 Excel: Microsoft® Excel® 2019 32bit 

- 開発テスト環境 OS: Windows 10 Pro

その他、実行環境など報告していただくと感謝感激雨霰。

# 使い方

## 「実行サンプル セル操作関係.xlsm」の使い方

「実行サンプル セル操作関係.xlsm」には「ModCell.bas」内のプロシージャの実行サンプルプロシージャのボタンが登録してある。

各ボタンを押して使用を確かめていただきたし
![実行サンプル中身](https://user-images.githubusercontent.com/73621859/130726860-90ccf952-910b-4212-8a4f-4a6a5406f25f.jpg)


## 設定

実行サンプル「実行サンプル セル操作関係.xlsm」の中の設定は以下の通り。

### 設定1（使用モジュール）

-  ModCell.bas
-  ModEnum.bas

### 設定2（参照ライブラリ）

特になし

## 現在「ModCell.bas」にて使用できるプロシージャ一覧

- GetBlankCell			…指定シート内の空白セルをオブジェクト形式で取得する
- SelectA1			…全シートのA1セルを選択した状態にする
- SortCell			…指定範囲のセルを並び替える
- SetCommentPicture		…セルのコメントで画像を表示する
- ResetFilter			…指定シートのフィルタを解除する。
- GetEndRow			…オートフィルタが設定してある場合も考慮しての最終行の取得
- GetEndCell			…オートフィルタが設定してある場合も考慮しての最終セル（オブジェクト形式）の取得
- SetCellDataBar		…0～1の値に基づいて、セルの書式設定でデータバーを設定する


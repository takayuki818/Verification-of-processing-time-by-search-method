Attribute VB_Name = "検証実行"
Option Explicit
Sub シートマッチ検索検証()
    Dim 始時 As Date, 検索値 As Long, 進捗 As Long
    Dim 検索結果(1 To 10000)
    始時 = Now
    実行中.Show vbModeless
    Call 実行中.プログレスバー更新(0)
    For 検索値 = 1 To 10000
        検索結果(検索値) = シートマッチ(検索値, "A:A")
        If Int(検索値 / 100) - 進捗 >= 10 Then
            進捗 = Int(検索値 / 100)
            Call 実行中.プログレスバー更新(進捗)
        End If
    Next
    Unload 実行中
    Sheets("検証").Cells(3, 3) = Int((Now - 始時) * 24 * 60 * 60 * 100) / 100
End Sub
Sub 配列マッチ検索検証()
    Dim 始時 As Date, 検索値 As Long, 進捗 As Long
    Dim 検索結果(1 To 10000)
    Dim 配列()
    始時 = Now
    実行中.Show vbModeless
    Call 実行中.プログレスバー更新(0)
    With Sheets("データ")
        配列 = Range("A1:A10000")
    End With
    For 検索値 = 1 To 10000
        検索結果(検索値) = 配列マッチ(検索値, 配列)
        If Int(検索値 / 100) - 進捗 >= 10 Then
            進捗 = Int(検索値 / 100)
            Call 実行中.プログレスバー更新(進捗)
        End If
    Next
    Unload 実行中
    Sheets("検証").Cells(4, 3) = Int((Now - 始時) * 24 * 60 * 60 * 100) / 100
End Sub
Sub シートネクスト検索検証()
    Dim 始時 As Date, 検索値 As Long, 終行 As Long, 進捗 As Long
    Dim 検索結果(1 To 10000)
    始時 = Now
    実行中.Show vbModeless
    Call 実行中.プログレスバー更新(0)
    終行 = Sheets("データ").Cells(Rows.Count, 1).End(xlUp).Row
    For 検索値 = 1 To 10000
        検索結果(検索値) = シートネクスト(検索値, 終行)
        If Int(検索値 / 100) - 進捗 >= 10 Then
            進捗 = Int(検索値 / 100)
            Call 実行中.プログレスバー更新(進捗)
        End If
    Next
    Unload 実行中
    Sheets("検証").Cells(5, 3) = Int((Now - 始時) * 24 * 60 * 60 * 100) / 100
End Sub
Sub 配列ネクスト検索検証()
    Dim 始時 As Date, 検索値 As Long, 進捗 As Long
    Dim 検索結果(1 To 10000)
    Dim 配列()
    始時 = Now
    実行中.Show vbModeless
    Call 実行中.プログレスバー更新(0)
    With Sheets("データ")
        配列 = Range("A1:A10000")
    End With
    For 検索値 = 1 To 10000
        検索結果(検索値) = 配列ネクスト(検索値, 配列)
        If Int(検索値 / 100) - 進捗 >= 10 Then
            進捗 = Int(検索値 / 100)
            Call 実行中.プログレスバー更新(進捗)
        End If
    Next
    Unload 実行中
    Sheets("検証").Cells(6, 3) = Int((Now - 始時) * 24 * 60 * 60 * 100) / 100
End Sub
Sub シートVLU検索検証()
    Dim 始時 As Date, 検索値 As Long, 進捗 As Long
    Dim 検索結果(1 To 10000)
    始時 = Now
    実行中.Show vbModeless
    Call 実行中.プログレスバー更新(0)
    For 検索値 = 1 To 10000
        検索結果(検索値) = シートVLU(検索値, "A:A")
        If Int(検索値 / 100) - 進捗 >= 10 Then
            進捗 = Int(検索値 / 100)
            Call 実行中.プログレスバー更新(進捗)
        End If
    Next
    Unload 実行中
    Sheets("検証").Cells(7, 3) = Int((Now - 始時) * 24 * 60 * 60 * 100) / 100
End Sub
Sub 配列VLU検索検証()
    Dim 始時 As Date, 検索値 As Long, 進捗 As Long
    Dim 検索結果(1 To 10000)
    Dim 配列()
    始時 = Now
    実行中.Show vbModeless
    Call 実行中.プログレスバー更新(0)
    With Sheets("データ")
        配列 = Range("A1:A10000")
    End With
    For 検索値 = 1 To 10000
        検索結果(検索値) = 配列VLU(検索値, 配列)
        If Int(検索値 / 100) - 進捗 >= 10 Then
            進捗 = Int(検索値 / 100)
            Call 実行中.プログレスバー更新(進捗)
        End If
    Next
    Unload 実行中
    Sheets("検証").Cells(8, 3) = Int((Now - 始時) * 24 * 60 * 60 * 100) / 100
End Sub

Attribute VB_Name = "検索FNC"
Option Explicit
Function シートマッチ(検索値, 範囲)
    With Sheets("データ")
        On Error Resume Next
        シートマッチ = WorksheetFunction.Match(検索値, .Range(範囲), 0)
    End With
End Function
Function 配列マッチ(検索値, 配列)
    With Sheets("データ")
        On Error Resume Next
        配列マッチ = WorksheetFunction.Match(検索値, 配列, 0)
    End With
End Function
Function シートネクスト(検索値, 終行)
    Dim 行 As Long
    With Sheets("データ")
        For 行 = 1 To 終行
            If .Cells(行, 1) = 検索値 Then
                シートネクスト = 行
                Exit For
            End If
        Next
    End With
End Function
Function 配列ネクスト(検索値, 配列)
    Dim 行 As Long
    With Sheets("データ")
        For 行 = 1 To UBound(配列, 1)
            If 配列(行, 1) = 検索値 Then
                配列ネクスト = 行
                Exit For
            End If
        Next
    End With
End Function
Function シートVLU(検索値, 範囲)
    With Sheets("データ")
        On Error Resume Next
        シートVLU = WorksheetFunction.VLookup(検索値, .Range(範囲), 1, 0)
    End With
End Function
Function 配列VLU(検索値, 配列)
    With Sheets("データ")
        On Error Resume Next
        配列VLU = WorksheetFunction.VLookup(検索値, 配列, 1, 0)
    End With
End Function

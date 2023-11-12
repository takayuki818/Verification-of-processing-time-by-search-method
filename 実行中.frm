VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} 実行中 
   Caption         =   "UserForm1"
   ClientHeight    =   888
   ClientLeft      =   48
   ClientTop       =   372
   ClientWidth     =   4284
   OleObjectBlob   =   "実行中.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "実行中"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub UserForm_Initialize()
    With Label1
        .Width = 0
    End With
End Sub
Sub プログレスバー更新(進捗 As Long)
    If 進捗 Mod 10 = 0 Then
        Label1.Width = 進捗 * 2
        Label2.Caption = 進捗 & "％完了"
        DoEvents
    End If
End Sub

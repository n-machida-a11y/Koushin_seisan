Attribute VB_Name = "ModError"
Option Explicit

' ============================================================
' 処理停止エラー
' 該当行を黄色ハイライト → ログ記録 → ポップアップ表示 → End で処理停止
' ws: 加工対象シート, rowNum: 問題のある行番号, message: 表示メッセージ
' ============================================================
Public Sub 処理停止エラー(ws As Worksheet, rowNum As Long, message As String)
    ' 該当行を黄色でハイライト
    ws.Rows(rowNum).Interior.Color = RGB(255, 255, 0)

    ' ログに記録
    Call ログ書込("エラー検出", "エラー", "行" & rowNum & ": " & message)

    ' ポップアップ表示
    MsgBox "【処理停止】" & vbCrLf & vbCrLf & _
           message & vbCrLf & vbCrLf & _
           "行番号: " & rowNum & vbCrLf & vbCrLf & _
           "オムロン担当者に問い合わせ後、データを修正して最初から再実行してください。", _
           vbCritical, "生産計画自動化 - 処理停止"

    ' 処理を終了（再開なし）
    End
End Sub

' ============================================================
' 警告ログ（続行）
' ログに記録するのみ。ポップアップなし・処理継続
' ============================================================
Public Sub 警告ログ(stepName As String, rowNum As Long, message As String)
    Call ログ書込(stepName, "警告", "行" & rowNum & ": " & message)
End Sub

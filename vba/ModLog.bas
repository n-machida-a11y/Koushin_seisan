Attribute VB_Name = "ModLog"
Option Explicit

' ============================================================
' ログシートに1行書き込む
' result: "成功" / "警告" / "エラー" / "情報"
' ============================================================
Public Sub ログ書込(stepName As String, result As String, message As String)
    Dim ws As Worksheet
    Dim nextRow As Long

    Set ws = ThisWorkbook.Sheets("ログ")
    nextRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1

    ws.Cells(nextRow, 1).Value = Now()
    ws.Cells(nextRow, 2).Value = stepName
    ws.Cells(nextRow, 3).Value = result
    ws.Cells(nextRow, 4).Value = message

    ' 警告・エラーは色でハイライト
    Select Case result
        Case "警告"
            ws.Cells(nextRow, 3).Interior.Color = RGB(255, 165, 0)
        Case "エラー"
            ws.Cells(nextRow, 3).Interior.Color = RGB(255, 100, 100)
    End Select
End Sub

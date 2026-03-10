Attribute VB_Name = "ModStep05"
Option Explicit

' ============================================================
' ステップ⑤: K列（追加仕様）に「計画生産対象」を含む行を削除する
' ============================================================
Public Sub Step05_計画生産対象削除(ws As Worksheet)
    Dim lastRow As Long
    Dim i As Long
    Dim cellVal As String
    Dim deletedCount As Long

    deletedCount = 0
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    ' 下から上に向かって削除（行削除時のインデックスズレを防ぐ）
    For i = lastRow To 2 Step -1
        cellVal = Trim(CStr(ws.Cells(i, g_ColTsuikashiyo).Value))
        If InStr(cellVal, "計画生産対象") > 0 Then
            ws.Rows(i).Delete
            deletedCount = deletedCount + 1
        End If
    Next i

    Call ログ書込("Step05_計画生産対象削除", "成功", deletedCount & "行を削除しました")
End Sub

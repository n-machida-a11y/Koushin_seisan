Attribute VB_Name = "ModStep09"
Option Explicit

' ============================================================
' ステップ⑨: 数量チェック（V8/V9 3ヶ月以内）
'
' MODEL（U列）が「V8」または「V9」（メンテ以外）で
' N列（出荷日）が3ヶ月以内の行に数量1以外があれば処理停止エラー
' ============================================================
Public Sub Step09_数量チェック(ws As Worksheet)
    Dim months3Later As Date
    months3Later = DateSerial(Year(g_BaseDate), Month(g_BaseDate) + 3, Day(g_BaseDate))

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    Dim i As Long
    For i = 2 To lastRow
        Dim model As String
        model = Trim(CStr(ws.Cells(i, g_ColModel).Value))

        ' V8またはV9（メンテ除く）のみチェック
        If model <> "V8" And model <> "V9" Then GoTo NextRow

        Dim shukkaDate As Variant
        shukkaDate = ws.Cells(i, g_ColShukkaDate).Value
        If IsEmpty(shukkaDate) Or CStr(shukkaDate) = "" Then GoTo NextRow
        If CDate(shukkaDate) > months3Later Then GoTo NextRow

        Dim suryo As Long
        suryo = CLng(ws.Cells(i, g_ColSuryo).Value)
        If suryo <> 1 Then
            Call 処理停止エラー(ws, i, _
                "MODEL「" & model & "」で数量が1ではありません（数量=" & suryo & "）。" & vbCrLf & _
                "オムロン担当者への問い合わせが必要です。" & vbCrLf & _
                "生産計画No: " & ws.Cells(i, g_ColSeisanNo).Value)
        End If
NextRow:
    Next i

    Call ログ書込("Step09_数量チェック", "成功", "V8/V9の3ヶ月以内数量チェック完了")
End Sub

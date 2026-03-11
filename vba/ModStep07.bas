Attribute VB_Name = "ModStep07"
Option Explicit

' ============================================================
' ステップ⑦: S列（BH型式TYPE）補完
' S列が空欄の行に型式を自動補完する
'
' 補完ロジック:
'   1. G列が「YB-」or「YU-」で始まる → G列の値をそのまま使用
'   2. それ以外 → 同一客先名（C列）の行からS列値を取得
'      G列も一致するものを優先
'
' 3ヶ月以内で補完不可 → 処理停止エラー（オムロン問い合わせ必要）
' 4ヶ月以降で補完不可 → 警告ログのみ（続行）
' ============================================================
Public Sub Step07_型式補完(ws As Worksheet)
    Dim lastRow As Long
    Dim i As Long
    Dim supplemented As Long
    Dim warned As Long

    supplemented = 0
    warned = 0
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    Dim months3Later As Date
    months3Later = DateSerial(Year(g_BaseDate), Month(g_BaseDate) + 3, Day(g_BaseDate))

    For i = g_DataStartRow To lastRow
        Dim bhType As String
        bhType = Trim(CStr(ws.Cells(i, g_ColBHType).Value))
        If bhType <> "" Then GoTo NextRow  ' すでに入力済みはスキップ

        Dim katashiki As String
        katashiki = Trim(CStr(ws.Cells(i, g_ColKatashiki).Value))

        ' 3ヶ月ルール判定
        Dim is3MonthsOrLess As Boolean
        is3MonthsOrLess = False
        Dim shukkaDate As Variant
        shukkaDate = ws.Cells(i, g_ColShukkaDate).Value
        If Not IsEmpty(shukkaDate) And CStr(shukkaDate) <> "" And IsDate(shukkaDate) Then
            is3MonthsOrLess = (CDate(shukkaDate) <= months3Later)
        End If

        ' パターン1: G列がYB-またはYU-で始まる → そのまま使用
        If Left(katashiki, 3) = "YB-" Or Left(katashiki, 3) = "YU-" Then
            ws.Cells(i, g_ColBHType).Value = katashiki
            supplemented = supplemented + 1
            GoTo NextRow
        End If

        ' パターン2: 客先名（C列）で同一行を検索して補完
        Dim foundType As String
        foundType = 客先名からBHType取得(ws, i, lastRow)

        If foundType <> "" Then
            ws.Cells(i, g_ColBHType).Value = foundType
            supplemented = supplemented + 1
        ElseIf is3MonthsOrLess Then
            ' 3ヶ月以内で補完不可 → 処理停止
            Call 処理停止エラー(ws, i, _
                "S列（BH型式TYPE）が空欄で自動補完できませんでした。" & vbCrLf & _
                "当月から3ヶ月以内のデータのため、オムロン担当者への問い合わせが必要です。" & vbCrLf & _
                "客先名: " & ws.Cells(i, g_ColKyakusakiName).Value)
        Else
            ' 4ヶ月以降で補完不可 → 警告のみで続行
            Call 警告ログ("Step07_型式補完", i, _
                "S列が空欄で補完不可（4ヶ月以降のデータのため続行）。客先名: " & _
                ws.Cells(i, g_ColKyakusakiName).Value)
            warned = warned + 1
        End If
NextRow:
    Next i

    Call ログ書込("Step07_型式補完", "成功", _
        supplemented & "行を補完、" & warned & "行を警告スキップ")
End Sub

' ============================================================
' 同一客先名（C列）の行からBH型式TYPE（S列）を取得する
' 複数候補がある場合はG列（型式）が同じものを優先する
' ============================================================
Private Function 客先名からBHType取得(ws As Worksheet, targetRow As Long, lastRow As Long) As String
    Dim targetKyakusaki As String
    Dim targetKatashiki As String
    targetKyakusaki = Trim(CStr(ws.Cells(targetRow, g_ColKyakusakiName).Value))
    targetKatashiki = Trim(CStr(ws.Cells(targetRow, g_ColKatashiki).Value))

    Dim exactMatch As String    ' G列も一致したもの（最優先）
    Dim partialMatch As String  ' 客先名だけ一致したもの
    exactMatch = ""
    partialMatch = ""

    Dim i As Long
    For i = g_DataStartRow To lastRow
        If i = targetRow Then GoTo NextRow

        Dim bhType As String
        bhType = Trim(CStr(ws.Cells(i, g_ColBHType).Value))
        If bhType = "" Then GoTo NextRow

        Dim kyakusaki As String
        kyakusaki = Trim(CStr(ws.Cells(i, g_ColKyakusakiName).Value))
        If kyakusaki <> targetKyakusaki Then GoTo NextRow

        ' 客先名一致: 最初のものをpartialMatchとして記録
        If partialMatch = "" Then partialMatch = bhType

        ' G列も一致するものが見つかればexactMatchとして記録し終了
        Dim katashiki As String
        katashiki = Trim(CStr(ws.Cells(i, g_ColKatashiki).Value))
        If katashiki = targetKatashiki Then
            exactMatch = bhType
            Exit For
        End If
NextRow:
    Next i

    If exactMatch <> "" Then
        客先名からBHType取得 = exactMatch
    Else
        客先名からBHType取得 = partialMatch
    End If
End Function

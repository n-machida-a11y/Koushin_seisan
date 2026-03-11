Attribute VB_Name = "ModStep08"
Option Explicit

' ============================================================
' ステップ⑧: 計画生産行展開（1台1行化）
'
' 対象: F列（機種名）に「計画生産」を含む行 かつ
'       N列（出荷日）が当月～3ヶ月以内
'
' 処理: L列（数量）の数だけ行をコピーして展開し、
'       B列（生産計画No）末尾に -01,-02... と連番を付与する
' ============================================================
Public Sub Step08_計画生産行展開(ws As Worksheet)
    Dim months3Later As Date
    Dim expandedCount As Long
    expandedCount = 0
    months3Later = DateSerial(Year(g_BaseDate), Month(g_BaseDate) + 3, Day(g_BaseDate))

    ' 下から処理することで行挿入後のインデックスズレを防ぐ
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    Dim i As Long
    For i = lastRow To g_DataStartRow Step -1
        Dim kishuName As String
        kishuName = Trim(CStr(ws.Cells(i, g_ColKishuName).Value))
        If InStr(kishuName, "計画生産") = 0 Then GoTo NextRow

        Dim shukkaDate As Variant
        shukkaDate = ws.Cells(i, g_ColShukkaDate).Value
        If IsEmpty(shukkaDate) Or CStr(shukkaDate) = "" Then GoTo NextRow
        If Not IsDate(shukkaDate) Then GoTo NextRow
        If CDate(shukkaDate) > months3Later Then GoTo NextRow

        Dim suryo As Long
        suryo = CLng(ws.Cells(i, g_ColSuryo).Value)
        If suryo <= 1 Then GoTo NextRow

        ' 元の生産計画Noを取得
        Dim baseNo As String
        baseNo = Trim(CStr(ws.Cells(i, g_ColSeisanNo).Value))

        ' 数量分の行をi+1以降に挿入してコピー（後ろから処理）
        Dim j As Long
        For j = suryo To 1 Step -1
            If j > 1 Then
                ws.Rows(i + 1).Insert Shift:=xlDown
                ws.Rows(i).Copy ws.Rows(i + 1)
            End If
            ws.Cells(i + j - 1, g_ColSeisanNo).Value = baseNo & "-" & Format(j, "00")
            ws.Cells(i + j - 1, g_ColSuryo).Value = 1
        Next j

        expandedCount = expandedCount + 1
NextRow:
    Next i

    Call ログ書込("Step08_計画生産行展開", "成功", expandedCount & "件の行展開を実施しました")
End Sub

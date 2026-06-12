Attribute VB_Name = "ModStep15"
Option Explicit

' ============================================================
' Step15: 順序指示出荷台数の入力
'
' BHPlan日程表から日付ごとのV8/V9出荷台数を集計し、
' 出荷・完了計画のE列(順序指示出荷台数)とG列(出荷台数LAZ計)に転記
' 合計行(SUM関数)の手前までを対象とする
' 当月以降で出荷が無くなった日は0クリア(減少更新)。過去の日付は保持
' ============================================================
Public Sub Step15_出荷台数入力(targetWs As Worksheet)
    Dim lastRow As Long
    lastRow = targetWs.Cells(targetWs.Rows.Count, 1).End(xlUp).Row
    
    Dim v8Counts As Object
    Set v8Counts = CreateObject("Scripting.Dictionary")
    Dim v9Counts As Object
    Set v9Counts = CreateObject("Scripting.Dictionary")
    
    Dim i As Long
    For i = g_DataStartRow To lastRow
        Dim model As String
        model = Trim(CStr(targetWs.Cells(i, g_ColModel).Value))
        If model <> "V8" And model <> "V9" Then GoTo NextRow
        
        Dim shukkaDate As Variant
        shukkaDate = targetWs.Cells(i, g_ColShukkaDate).Value
        If IsEmpty(shukkaDate) Or Not IsDate(shukkaDate) Then GoTo NextRow
        If CDate(shukkaDate) < g_BaseDate Then GoTo NextRow
        
        Dim suryo As Variant
        suryo = targetWs.Cells(i, g_ColSuryo).Value
        If IsEmpty(suryo) Or Not IsNumeric(suryo) Then GoTo NextRow
        
        Dim dateKey As String
        dateKey = Format(CDate(shukkaDate), "YYYY/MM/DD")
        
        If model = "V8" Then
            If v8Counts.Exists(dateKey) Then
                v8Counts(dateKey) = v8Counts(dateKey) + CLng(suryo)
            Else
                v8Counts.Add dateKey, CLng(suryo)
            End If
        Else
            If v9Counts.Exists(dateKey) Then
                v9Counts(dateKey) = v9Counts(dateKey) + CLng(suryo)
            Else
                v9Counts.Add dateKey, CLng(suryo)
            End If
        End If
NextRow:
    Next i
    
    Dim v8Written As Long
    v8Written = 0
    If g_V8ProdSchedulePath <> "" Then
        v8Written = 台数書込(g_V8ProdSchedulePath, g_SheetV8ShukkaKeikaku, v8Counts)
    End If
    
    Dim v9Written As Long
    v9Written = 0
    If g_V9ProdSchedulePath <> "" Then
        v9Written = 台数書込(g_V9ProdSchedulePath, g_SheetV9ShukkaKeikaku, v9Counts)
    End If
    
    Call ログ書込("Step15_出荷台数入力", "完了", _
        "V8: " & v8Written & "日分、V9: " & v9Written & "日分の出荷台数を入力")
End Sub

' ============================================================
' 出荷・完了計画シートのE列/G列に台数を書き込む
' 合計行(SUM関数)の手前までを対象
' ============================================================
Private Function 台数書込(filePath As String, sheetName As String, counts As Object) As Long
    Dim wb As Workbook
    Dim ws As Worksheet
    
    On Error Resume Next
    Set wb = Workbooks.Open(filePath)
    On Error GoTo 0
    If wb Is Nothing Then
        台数書込 = 0
        Exit Function
    End If
    
    Set ws = シート検索(wb, sheetName)
    If ws Is Nothing Then
        wb.Close SaveChanges:=False
        台数書込 = 0
        Exit Function
    End If
    
    ' 合計行を探す（B列が空でE列とF列とG列に値がある行）
    Dim sumRow As Long
    sumRow = 0
    Dim scanLast As Long
    scanLast = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row
    
    Dim r As Long
    For r = 6 To scanLast + 20  ' 合計行はB列最終行の直後にあるため余裕を持つ
        Dim bEmpty As Boolean
        bEmpty = IsEmpty(ws.Cells(r, 2).Value) Or Trim(CStr(ws.Cells(r, 2).Value)) = ""
        If bEmpty Then
            If Not IsEmpty(ws.Cells(r, 5).Value) And _
               Not IsEmpty(ws.Cells(r, 6).Value) And _
               Not IsEmpty(ws.Cells(r, 7).Value) Then
                sumRow = r
                Exit For
            End If
        End If
    Next r
    
    ' 合計行が見つからなければ全行を対象
    Dim endRow As Long
    If sumRow > 0 Then
        endRow = sumRow - 1
    Else
        endRow = scanLast
    End If
    
    Dim writtenCount As Long
    writtenCount = 0
    
    For r = 6 To endRow
        Dim cellDate As Variant
        cellDate = ws.Cells(r, 2).Value
        If Not IsDate(cellDate) Then GoTo NextDateRow
        
        Dim dateKey As String
        dateKey = Format(CDate(cellDate), "YYYY/MM/DD")
        
        If counts.Exists(dateKey) Then
            ws.Cells(r, 5).Value = counts(dateKey)
            ws.Cells(r, 7).Value = counts(dateKey)
            writtenCount = writtenCount + 1
        ElseIf CDate(cellDate) >= g_BaseDate Then
            ' 当月以降で出荷が無くなった日は0に更新(減少更新の反映)。
            ' 稼働日(○)は0、非稼働日は空欄(手作業の正解Excelに合わせる)
            If Trim(CStr(ws.Cells(r, 4).Value)) = "○" Then
                ws.Cells(r, 5).Value = 0
                ws.Cells(r, 7).Value = 0
            Else
                ws.Cells(r, 5).ClearContents
                ws.Cells(r, 7).ClearContents
            End If
        End If
NextDateRow:
    Next r
    
    wb.Save
    wb.Close SaveChanges:=False
    
    台数書込 = writtenCount
End Function

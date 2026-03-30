Attribute VB_Name = "ModStep14"
Option Explicit

' ============================================================
' Step14: 出荷・完了計画に当月+3ヶ月後の日付を追加
'
' V8/V9のProduction Scheduleの「出荷・完了計画」シートの
' 合計行(SUM関数がある行)の上に、まだない日付行を挿入する
' D列に稼働日フラグ(○/×)を設定
' ============================================================
Public Sub Step14_出荷完了計画日付追加(targetWs As Worksheet)
    Dim months3Later As Date
    months3Later = DateSerial(Year(g_BaseDate), Month(g_BaseDate) + 4, 0)
    
    Dim addedV8 As Long
    Dim addedV9 As Long
    addedV8 = 0
    addedV9 = 0
    
    If g_V8ProdSchedulePath <> "" Then
        addedV8 = 日付追加処理(g_V8ProdSchedulePath, g_SheetV8ShukkaKeikaku, months3Later)
    End If
    
    If g_V9ProdSchedulePath <> "" Then
        addedV9 = 日付追加処理(g_V9ProdSchedulePath, g_SheetV9ShukkaKeikaku, months3Later)
    End If
    
    Call ログ書込("Step14_出荷完了計画日付追加", "完了", _
        "V8: " & addedV8 & "日追加、V9: " & addedV9 & "日追加")
End Sub

' ============================================================
' 指定ファイルの出荷・完了計画シートに日付行を追加
' 合計行(E列にSUM関数がある行)の上に挿入する
' ============================================================
Private Function 日付追加処理(filePath As String, sheetName As String, endDate As Date) As Long
    Dim wb As Workbook
    Dim ws As Worksheet
    
    On Error Resume Next
    Set wb = Workbooks.Open(filePath)
    On Error GoTo 0
    If wb Is Nothing Then
        Call ログ書込("Step14", "警告", "ファイルを開けません: " & filePath)
        日付追加処理 = 0
        Exit Function
    End If
    
    Set ws = シート検索(wb, sheetName)
    If ws Is Nothing Then
        Call ログ書込("Step14", "警告", "シートが見つかりません: " & sheetName)
        wb.Close SaveChanges:=False
        日付追加処理 = 0
        Exit Function
    End If
    
    ' 合計行を探す（B列が空でE列とF列とG列に値がある行）
    Dim sumRow As Long
    sumRow = 0
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row
    
    Dim r As Long
    For r = 6 To lastRow
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
    
    If sumRow = 0 Then
        Call ログ書込("Step14", "警告", "合計行(SUM関数)が見つかりません: " & sheetName)
        wb.Close SaveChanges:=False
        日付追加処理 = 0
        Exit Function
    End If
    
    ' 既存の日付を収集（B列、合計行の手前まで）
    Dim existingDates As Object
    Set existingDates = CreateObject("Scripting.Dictionary")
    For r = 6 To sumRow - 1
        Dim cellVal As Variant
        cellVal = ws.Cells(r, 2).Value
        If IsDate(cellVal) Then
            Dim dateKey As String
            dateKey = Format(CDate(cellVal), "YYYY/MM/DD")
            If Not existingDates.Exists(dateKey) Then
                existingDates.Add dateKey, r
            End If
        End If
    Next r
    
    ' 当月1日からendDateまで、ない日付を合計行の上に挿入
    Dim currentDate As Date
    Dim addedCount As Long
    addedCount = 0
    currentDate = g_BaseDate
    
    Do While currentDate <= endDate
        dateKey = Format(currentDate, "YYYY/MM/DD")
        
        If Not existingDates.Exists(dateKey) Then
            ' 合計行の上に行を挿入
            ws.Rows(sumRow).Insert Shift:=xlDown
            
            ' 挿入した行にデータを設定
            ' A列: 年度（4月始まり。1～3月は前年度）
            Dim nendo As Long
            If Month(currentDate) >= 4 Then
                nendo = Year(currentDate) Mod 100
            Else
                nendo = (Year(currentDate) - 1) Mod 100
            End If
            ws.Cells(sumRow, 1).Value = nendo
            ' B列: MM/DD形式の日付
            ws.Cells(sumRow, 2).Value = currentDate
            ws.Cells(sumRow, 2).NumberFormat = "M/D"
            ws.Cells(sumRow, 3).Value = Mid("月火水木金土日", Weekday(currentDate, vbMonday), 1)
            
            ' D列: 稼働日フラグ
            Dim wd As Long
            wd = Weekday(currentDate, vbMonday)
            If wd >= 6 Then
                ws.Cells(sumRow, 4).Value = "×"
            Else
                ws.Cells(sumRow, 4).Value = "○"
            End If
            
            ' 合計行が1行下にずれるので更新
            sumRow = sumRow + 1
            addedCount = addedCount + 1
        End If
        
        currentDate = currentDate + 1
    Loop
    
    wb.Save
    wb.Close SaveChanges:=False
    
    日付追加処理 = addedCount
End Function

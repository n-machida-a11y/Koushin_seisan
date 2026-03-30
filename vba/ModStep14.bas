Attribute VB_Name = "ModStep14"
Option Explicit

' ============================================================
' ステップ⑭: 出荷・完了計画に当月+3ヶ月後の日付を追加
'
' V8/V9のProduction Scheduleの「出荷・完了計画」シートに
' 当月から3ヶ月後までの日付行を追加する（まだない日付のみ）
' D列に稼働日フラグ(○/×)を設定
' ============================================================
Public Sub Step14_出荷完了計画日付追加(targetWs As Worksheet)
    Dim months3Later As Date
    months3Later = DateSerial(Year(g_BaseDate), Month(g_BaseDate) + 4, 0)  ' 3ヶ月後の末日
    
    Dim addedV8 As Long
    Dim addedV9 As Long
    addedV8 = 0
    addedV9 = 0
    
    ' V8 Production Schedule
    If g_V8ProdSchedulePath <> "" Then
        addedV8 = 日付追加処理(g_V8ProdSchedulePath, g_SheetV8ShukkaKeikaku, months3Later)
    End If
    
    ' V9 Production Schedule
    If g_V9ProdSchedulePath <> "" Then
        addedV9 = 日付追加処理(g_V9ProdSchedulePath, g_SheetV9ShukkaKeikaku, months3Later)
    End If
    
    Call ログ書込("Step14_出荷完了計画日付追加", "完了", _
        "V8: " & addedV8 & "日追加、V9: " & addedV9 & "日追加")
End Sub

' ============================================================
' 指定ファイルの出荷・完了計画シートに日付行を追加
' 戻り値: 追加した日数
' ============================================================
Private Function 日付追加処理(filePath As String, sheetName As String, endDate As Date) As Long
    Dim wb As Workbook
    Dim ws As Worksheet
    
    ' ファイルを開く
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
    
    ' 既存の日付を収集（B列）
    Dim existingDates As Object
    Set existingDates = CreateObject("Scripting.Dictionary")
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row
    
    Dim r As Long
    For r = 2 To lastRow
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
    
    ' 当月1日から endDate まで、日ごとにチェック
    Dim currentDate As Date
    Dim addedCount As Long
    addedCount = 0
    currentDate = g_BaseDate
    
    Do While currentDate <= endDate
        dateKey = Format(currentDate, "YYYY/MM/DD")
        
        If Not existingDates.Exists(dateKey) Then
            ' 新しい行を追加
            lastRow = lastRow + 1
            ws.Cells(lastRow, 1).Value = Format(currentDate, "YY/MM")  ' A列: 年月
            ws.Cells(lastRow, 2).Value = currentDate                    ' B列: 日付
            ws.Cells(lastRow, 3).Value = WeekdayName(Weekday(currentDate), True)  ' C列: 曜日
            
            ' D列: 稼働日フラグ（土日は×、平日は○。祝日は稼働日カレンダーで別途対応）
            Dim wd As Long
            wd = Weekday(currentDate, vbMonday)  ' 月=1, 日=7
            If wd >= 6 Then  ' 土日
                ws.Cells(lastRow, 4).Value = "×"
            Else
                ws.Cells(lastRow, 4).Value = "○"
            End If
            
            addedCount = addedCount + 1
        End If
        
        currentDate = currentDate + 1
    Loop
    
    ' ※結合セルがあるためソートは行わない（日付は末尾に追加される）
    
    wb.Save
    wb.Close SaveChanges:=False
    
    日付追加処理 = addedCount
End Function

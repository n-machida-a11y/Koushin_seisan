Attribute VB_Name = "ModStep15"
Option Explicit

' ============================================================
' ステップ⑮: 順序指示出荷台数の入力
'
' BHPlan日程表の集計データから日付ごとの出荷台数を取得し、
' 出荷・完了計画のE列(順序指示出荷台数)とG列(出荷台数LAZ計)に転記
' ============================================================
Public Sub Step15_出荷台数入力(targetWs As Worksheet)
    Dim lastRow As Long
    lastRow = targetWs.Cells(targetWs.Rows.Count, 1).End(xlUp).Row
    
    ' --- BHPlanから日付ごとの出荷台数を集計 ---
    ' V8用: date -> 台数
    Dim v8Counts As Object
    Set v8Counts = CreateObject("Scripting.Dictionary")
    ' V9用: date -> 台数
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
    
    ' --- V8出荷・完了計画に書き込み ---
    Dim v8Written As Long
    v8Written = 0
    If g_V8ProdSchedulePath <> "" And v8Counts.Count > 0 Then
        v8Written = 台数書込(g_V8ProdSchedulePath, "Ｖ８出荷・完了計画", v8Counts)
    End If
    
    ' --- V9出荷・完了計画に書き込み ---
    Dim v9Written As Long
    v9Written = 0
    If g_V9ProdSchedulePath <> "" And v9Counts.Count > 0 Then
        v9Written = 台数書込(g_V9ProdSchedulePath, "Ｖ９－BH出荷・完了計画", v9Counts)
    End If
    
    Call ログ書込("Step15_出荷台数入力", "完了", _
        "V8: " & v8Written & "日分、V9: " & v9Written & "日分の出荷台数を入力")
End Sub

' ============================================================
' 出荷・完了計画シートのE列/G列に台数を書き込む
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
    
    On Error Resume Next
    Set ws = wb.Sheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then
        wb.Close SaveChanges:=False
        台数書込 = 0
        Exit Function
    End If
    
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row
    
    Dim writtenCount As Long
    writtenCount = 0
    
    Dim r As Long
    For r = 6 To lastRow
        Dim cellDate As Variant
        cellDate = ws.Cells(r, 2).Value
        If Not IsDate(cellDate) Then GoTo NextDateRow
        
        Dim dateKey As String
        dateKey = Format(CDate(cellDate), "YYYY/MM/DD")
        
        If counts.Exists(dateKey) Then
            ws.Cells(r, 5).Value = counts(dateKey)  ' E列: 順序指示出荷台数
            ws.Cells(r, 7).Value = counts(dateKey)  ' G列: 出荷台数LAZ計
            writtenCount = writtenCount + 1
        End If
NextDateRow:
    Next r
    
    wb.Save
    wb.Close SaveChanges:=False
    
    台数書込 = writtenCount
End Function

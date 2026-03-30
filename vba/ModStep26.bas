Attribute VB_Name = "ModStep26"
Option Explicit

' ============================================================
' Step26: Forecastデータ完成
'
' BHPlan日程表のshipment month(V列)をキーに
' BP列以降のコンポーネント台数を集計
' ============================================================
Public Sub Step26_Forecast完成(targetWs As Worksheet)
    Dim lastRow As Long
    lastRow = targetWs.Cells(targetWs.Rows.Count, 1).End(xlUp).Row
    Dim headerRow As Long
    headerRow = g_DataStartRow - 1
    
    Dim bpStartCol As Long
    bpStartCol = 0
    Dim c As Long
    For c = 60 To 100
        Dim hVal As String
        hVal = Trim(CStr(targetWs.Cells(headerRow, c).Value))
        If InStr(hVal, "LAA101Q") > 0 Or InStr(hVal, "LAA101") > 0 Then
            bpStartCol = c
            Exit For
        End If
    Next c
    
    If bpStartCol = 0 Then
        Call ログ書込("Step26_Forecast完成", "警告", "LAA101Q列が見つかりません")
        Exit Sub
    End If
    
    Dim compCount As Long
    compCount = 0
    Dim lastCol As Long
    lastCol = targetWs.UsedRange.Columns.Count
    For c = bpStartCol To lastCol
        If Trim(CStr(targetWs.Cells(headerRow, c).Value)) <> "" Then compCount = compCount + 1
    Next c
    
    Dim forecast As Object
    Set forecast = CreateObject("Scripting.Dictionary")
    Dim monthList As Object
    Set monthList = CreateObject("Scripting.Dictionary")
    
    Dim i As Long
    For i = g_DataStartRow To lastRow
        Dim shipMonth As Variant
        shipMonth = targetWs.Cells(i, g_ColShipmentMonth).Value
        If IsEmpty(shipMonth) Then GoTo NextRow
        Dim mk As String
        mk = Trim(CStr(shipMonth))
        If mk = "" Then GoTo NextRow
        If Not monthList.Exists(mk) Then monthList.Add mk, 0
        For c = bpStartCol To lastCol
            Dim qty As Variant
            qty = targetWs.Cells(i, c).Value
            If IsNumeric(qty) And Not IsEmpty(qty) Then
                If CLng(qty) > 0 Then
                    Dim fk As String
                    fk = mk & "|" & c
                    If forecast.Exists(fk) Then
                        forecast(fk) = forecast(fk) + CLng(qty)
                    Else
                        forecast.Add fk, CLng(qty)
                    End If
                End If
            End If
        Next c
NextRow:
    Next i
    
    Call ログ書込("Step26_Forecast完成", "完了", _
        "Forecastデータ集計完了（" & monthList.Count & "ヶ月、" & compCount & "コンポーネント、" & forecast.Count & "件）")
End Sub

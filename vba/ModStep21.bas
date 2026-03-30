Attribute VB_Name = "ModStep21"
Option Explicit

' ============================================================
' Step21: KMPユニット日当たり必要台数
'
' 星取表のコンポーネント列の「・」マークから日ごとの必要台数を集計
' ============================================================
Public Sub Step21_KMPユニット必要台数(targetWs As Worksheet)
    If g_V8ProdSchedulePath = "" Then
        Call ログ書込("Step21_KMPユニット必要台数", "警告", "V8_ProductionScheduleパスが未設定です")
        Exit Sub
    End If
    
    Dim psWb As Workbook
    Set psWb = Workbooks.Open(g_V8ProdSchedulePath)
    
    Dim hsWs As Worksheet
    Set hsWs = シート検索(psWb, g_SheetV8Hoshitori)
    If hsWs Is Nothing Then
        Call ログ書込("Step21", "エラー", "V8星取表シートが見つかりません（設定: " & g_SheetV8Hoshitori & "）")
        psWb.Close SaveChanges:=False
        Exit Sub
    End If
    
    Dim lastRow As Long
    lastRow = hsWs.Cells(hsWs.Rows.Count, 1).End(xlUp).Row
    
    Dim dailyCounts As Object
    Set dailyCounts = CreateObject("Scripting.Dictionary")
    Dim totalUnits As Long
    totalUnits = 0
    
    Dim r As Long
    For r = 7 To lastRow
        Dim dateO As Variant
        Dim dateP As Variant
        dateO = hsWs.Cells(r, 15).Value
        dateP = hsWs.Cells(r, 16).Value
        
        Dim c As Long
        For c = 19 To 41
            Dim cellVal As String
            cellVal = Trim(CStr(hsWs.Cells(r, c).Value))
            If cellVal = "・" Or cellVal = "･" Then
                totalUnits = totalUnits + 1
                Dim useDate As Variant
                If c <= 25 Then
                    useDate = dateO
                Else
                    useDate = dateP
                End If
                If IsDate(useDate) Then
                    Dim dk As String
                    dk = Format(CDate(useDate), "YYYY/MM/DD")
                    If dailyCounts.Exists(dk) Then
                        dailyCounts(dk) = dailyCounts(dk) + 1
                    Else
                        dailyCounts.Add dk, 1
                    End If
                End If
            End If
        Next c
    Next r
    
    psWb.Close SaveChanges:=False
    
    Call ログ書込("Step21_KMPユニット必要台数", "完了", _
        "KMPユニット必要台数を集計（合計: " & totalUnits & "ユニット、" & dailyCounts.Count & "日分）")
End Sub

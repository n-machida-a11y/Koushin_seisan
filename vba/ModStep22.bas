Attribute VB_Name = "ModStep22"
Option Explicit

' ============================================================
' Step22: KMP SHIPMENT PLAN更新
' ============================================================
Public Sub Step22_KMPシップメント更新(targetWs As Worksheet)
    If g_V8ProdSchedulePath = "" Then
        Call ログ書込("Step22_KMPシップメント更新", "警告", "V8_ProductionScheduleパスが未設定です")
        Exit Sub
    End If
    
    Dim psWb As Workbook
    Set psWb = Workbooks.Open(g_V8ProdSchedulePath)
    
    Dim kmpWs As Worksheet
    Set kmpWs = シート検索(psWb, g_SheetV8KMPShipment)
    If kmpWs Is Nothing Then
        Call ログ書込("Step22", "エラー", "V8 KMP SHIPMENT PLANシートが見つかりません（設定: " & g_SheetV8KMPShipment & "）")
        psWb.Close SaveChanges:=False
        Exit Sub
    End If
    
    Dim lastRow As Long
    lastRow = targetWs.Cells(targetWs.Rows.Count, 1).End(xlUp).Row
    
    Dim dateCounts As Object
    Set dateCounts = CreateObject("Scripting.Dictionary")
    
    Dim i As Long
    For i = g_DataStartRow To lastRow
        Dim model As String
        model = Trim(CStr(targetWs.Cells(i, g_ColModel).Value))
        If model <> "V8" Then GoTo NextRow
        
        Dim shukkaDate As Variant
        shukkaDate = targetWs.Cells(i, g_ColShukkaDate).Value
        If IsEmpty(shukkaDate) Or Not IsDate(shukkaDate) Then GoTo NextRow
        If CDate(shukkaDate) < g_BaseDate Then GoTo NextRow
        
        Dim suryo As Variant
        suryo = targetWs.Cells(i, g_ColSuryo).Value
        If IsEmpty(suryo) Or Not IsNumeric(suryo) Then GoTo NextRow
        
        Dim dk As String
        dk = Format(CDate(shukkaDate), "YYYY/MM/DD")
        If dateCounts.Exists(dk) Then
            dateCounts(dk) = dateCounts(dk) + CLng(suryo)
        Else
            dateCounts.Add dk, CLng(suryo)
        End If
NextRow:
    Next i
    
    Dim kmpLastRow As Long
    kmpLastRow = kmpWs.Cells(kmpWs.Rows.Count, 1).End(xlUp).Row
    Dim writtenCount As Long
    writtenCount = 0
    Dim r As Long
    For r = 7 To kmpLastRow
        Dim cellDate As Variant
        cellDate = kmpWs.Cells(r, 1).Value
        If Not IsDate(cellDate) Then GoTo NextKMPRow
        Dim dateKey As String
        dateKey = Format(CDate(cellDate), "YYYY/MM/DD")
        If dateCounts.Exists(dateKey) Then
            kmpWs.Cells(r, 29).Value = dateCounts(dateKey)
            writtenCount = writtenCount + 1
        End If
NextKMPRow:
    Next r
    
    psWb.Save
    psWb.Close SaveChanges:=False
    Call ログ書込("Step22_KMPシップメント更新", "完了", writtenCount & "日分のKMP SHIPMENT PLANを更新しました")
End Sub

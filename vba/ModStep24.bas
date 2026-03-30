Attribute VB_Name = "ModStep24"
Option Explicit

' ============================================================
' Step24: KMP出荷スケジュール更新
' ============================================================
Public Sub Step24_KMP出荷スケジュール更新(targetWs As Worksheet)
    If g_V8ProdSchedulePath = "" Then
        Call ログ書込("Step24_KMP出荷スケジュール更新", "警告", "V8_ProductionScheduleパスが未設定です")
        Exit Sub
    End If
    
    Dim psWb As Workbook
    Set psWb = Workbooks.Open(g_V8ProdSchedulePath)
    
    Dim kmpWs As Worksheet
    Set kmpWs = シート検索(psWb, g_SheetV8KMPShipment)
    If kmpWs Is Nothing Then
        Call ログ書込("Step24", "エラー", "KMP SHIPMENT PLANシートが見つかりません")
        psWb.Close SaveChanges:=False
        Exit Sub
    End If
    
    Dim shipmentData As Object
    Set shipmentData = CreateObject("Scripting.Dictionary")
    Dim kmpLastRow As Long
    kmpLastRow = kmpWs.Cells(kmpWs.Rows.Count, 1).End(xlUp).Row
    Dim r As Long
    For r = 7 To kmpLastRow
        Dim kmpDate As Variant
        kmpDate = kmpWs.Cells(r, 1).Value
        If IsDate(kmpDate) Then
            Dim total As Variant
            total = kmpWs.Cells(r, 29).Value
            If IsNumeric(total) And Not IsEmpty(total) Then
                shipmentData(Format(CDate(kmpDate), "YYYY/MM/DD")) = CLng(total)
            End If
        End If
    Next r
    
    Dim schedWs As Worksheet
    Set schedWs = シート検索(psWb, g_SheetKMPSchedule)
    If schedWs Is Nothing Then
        Call ログ書込("Step24", "エラー", "KMP出荷スケジュールシートが見つかりません（設定: " & g_SheetKMPSchedule & "）")
        psWb.Close SaveChanges:=False
        Exit Sub
    End If
    
    Dim schedLastRow As Long
    schedLastRow = schedWs.Cells(schedWs.Rows.Count, 1).End(xlUp).Row
    Dim writtenCount As Long
    writtenCount = 0
    For r = 2 To schedLastRow
        Dim schedDate As Variant
        schedDate = schedWs.Cells(r, 1).Value
        If IsDate(schedDate) Then
            Dim dk As String
            dk = Format(CDate(schedDate), "YYYY/MM/DD")
            If shipmentData.Exists(dk) Then
                schedWs.Cells(r, 12).Value = shipmentData(dk)
                schedWs.Cells(r, 13).Value = shipmentData(dk)
                writtenCount = writtenCount + 1
            End If
        End If
    Next r
    
    psWb.Save
    psWb.Close SaveChanges:=False
    Call ログ書込("Step24_KMP出荷スケジュール更新", "完了", writtenCount & "日分のKMP出荷スケジュールを更新しました")
End Sub

Attribute VB_Name = "ModStep25"
Option Explicit

' ============================================================
' Step25: KMP MPS更新
' ============================================================
Public Sub Step25_KMP_MPS更新(targetWs As Worksheet)
    If g_V8ProdSchedulePath = "" Then
        Call ログ書込("Step25_KMP_MPS更新", "警告", "V8_ProductionScheduleパスが未設定です")
        Exit Sub
    End If
    
    Dim psWb As Workbook
    Set psWb = Workbooks.Open(g_V8ProdSchedulePath)
    
    Dim mpsWs As Worksheet
    Set mpsWs = シート検索(psWb, g_SheetKMPMPS)
    If mpsWs Is Nothing Then
        Call ログ書込("Step25", "エラー", "KMP MPSシートが見つかりません（設定: " & g_SheetKMPMPS & "）")
        psWb.Close SaveChanges:=False
        Exit Sub
    End If
    
    Dim lastRow As Long
    lastRow = targetWs.Cells(targetWs.Rows.Count, 1).End(xlUp).Row
    Dim monthlyCounts As Object
    Set monthlyCounts = CreateObject("Scripting.Dictionary")
    
    Dim i As Long
    For i = g_DataStartRow To lastRow
        Dim model As String
        model = Trim(CStr(targetWs.Cells(i, g_ColModel).Value))
        If model <> "V8" And model <> "V9" Then GoTo NextRow
        Dim shipMonth As Variant
        shipMonth = targetWs.Cells(i, g_ColShipmentMonth).Value
        If IsEmpty(shipMonth) Then GoTo NextRow
        Dim suryo As Variant
        suryo = targetWs.Cells(i, g_ColSuryo).Value
        If IsEmpty(suryo) Or Not IsNumeric(suryo) Then GoTo NextRow
        Dim mk As String
        mk = Trim(CStr(shipMonth))
        If mk = "" Then GoTo NextRow
        If monthlyCounts.Exists(mk) Then
            monthlyCounts(mk) = monthlyCounts(mk) + CLng(suryo)
        Else
            monthlyCounts.Add mk, CLng(suryo)
        End If
NextRow:
    Next i
    
    Dim mpsLastRow As Long
    mpsLastRow = mpsWs.Cells(mpsWs.Rows.Count, 1).End(xlUp).Row
    Dim writtenCount As Long
    writtenCount = 0
    Dim mk2 As Variant
    For Each mk2 In monthlyCounts.Keys
        Dim found As Boolean
        found = False
        Dim r As Long
        For r = 2 To mpsLastRow
            If Trim(CStr(mpsWs.Cells(r, 1).Value)) = CStr(mk2) Then
                mpsWs.Cells(r, 2).Value = monthlyCounts(mk2)
                found = True
                writtenCount = writtenCount + 1
                Exit For
            End If
        Next r
        If Not found Then
            mpsLastRow = mpsLastRow + 1
            mpsWs.Cells(mpsLastRow, 1).Value = CStr(mk2)
            mpsWs.Cells(mpsLastRow, 2).Value = monthlyCounts(mk2)
            writtenCount = writtenCount + 1
        End If
    Next mk2
    
    psWb.Save
    psWb.Close SaveChanges:=False
    Call ログ書込("Step25_KMP_MPS更新", "完了", writtenCount & "ヶ月分のKMP MPSデータを更新しました")
End Sub

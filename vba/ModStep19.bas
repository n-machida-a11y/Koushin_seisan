Attribute VB_Name = "ModStep19"
Option Explicit

' ============================================================
' Step19: 星取表の差し替え
'
' Production Scheduleの星取表シートの情報を最新データに差し替える。
'
' 差し替え起点: 来月1日（日程変更があればそこから）
' スキップ条件:
'   - R列に日付が入っている行 → スキップ
'   - N列に「M」が入っている行 → スキップ
'   - S列～AY列に日付が入っている行 → スキップ（ログ記録）
' ============================================================
Public Sub Step19_星取表差替(targetWs As Worksheet)
    Dim replacedV8 As Long
    Dim replacedV9 As Long
    replacedV8 = 0
    replacedV9 = 0
    
    ' V8星取表の差し替え
    If g_V8ProdSchedulePath <> "" Then
        replacedV8 = 星取表差替処理(targetWs, "V8")
    End If
    
    ' V9星取表の差し替え
    If g_V9ProdSchedulePath <> "" Then
        replacedV9 = 星取表差替処理(targetWs, "V9")
    End If
    
    Call ログ書込("Step19_星取表差替", "完了", _
        "V8: " & replacedV8 & "行差替、V9: " & replacedV9 & "行差替")
End Sub

' ============================================================
' 指定MODELの星取表差し替え処理
' ============================================================
Private Function 星取表差替処理(targetWs As Worksheet, modelType As String) As Long
    ' Production Scheduleを開く
    Dim psPath As String
    Dim hsSheetName As String
    If modelType = "V8" Then
        psPath = g_V8ProdSchedulePath
        hsSheetName = g_SheetV8Hoshitori
    Else
        psPath = g_V9ProdSchedulePath
        hsSheetName = g_SheetV9Hoshitori
    End If
    
    Dim psWb As Workbook
    Set psWb = Workbooks.Open(psPath)
    Dim hsWs As Worksheet
    Set hsWs = シート検索(psWb, hsSheetName)
    
    If hsWs Is Nothing Then
        Call ログ書込("Step19", "警告", modelType & "星取表シートが見つかりません")
        psWb.Close SaveChanges:=False
        星取表差替処理 = 0
        Exit Function
    End If
    
    ' KP-No列の特定
    Dim hsKPCol As Long
    If modelType = "V8" Then
        hsKPCol = g_V8SavedKPNoCol
    Else
        hsKPCol = g_V9SavedKPNoCol
    End If
    
    ' 星取表の出荷日列
    Dim hsDateCol As Long
    If modelType = "V8" Then
        hsDateCol = 11  ' K列
    Else
        hsDateCol = 8   ' H列
    End If
    
    ' 差し替え起点の決定
    ' 基本は来月1日。日程変更があればそこから。
    Dim nextMonth As Date
    nextMonth = DateSerial(Year(g_BaseDate), Month(g_BaseDate) + 1, 1)
    Dim replaceFrom As Date
    replaceFrom = nextMonth
    
    ' BHPlanから対象MODELのデータを収集（KP-No→出荷日マップ）
    Dim bhpDates As Object
    Set bhpDates = CreateObject("Scripting.Dictionary")
    Dim bhpLastRow As Long
    bhpLastRow = targetWs.Cells(targetWs.Rows.Count, 1).End(xlUp).Row
    
    Dim i As Long
    For i = g_DataStartRow To bhpLastRow
        Dim model As String
        model = Trim(CStr(targetWs.Cells(i, g_ColModel).Value))
        If model <> modelType Then GoTo NextBHPRow
        
        Dim kp As String
        kp = Trim(CStr(targetWs.Cells(i, g_ColKPNo).Value))
        If kp = "" Then GoTo NextBHPRow
        
        Dim dt As Variant
        dt = targetWs.Cells(i, g_ColShukkaDate).Value
        If IsDate(dt) Then
            bhpDates(kp) = CDate(dt)
        End If
NextBHPRow:
    Next i
    
    ' 日程変更の検出（星取表の出荷日とBHPlanの出荷日が異なる行を探す）
    Dim hsLastRow As Long
    hsLastRow = hsWs.Cells(hsWs.Rows.Count, 1).End(xlUp).Row
    
    Dim r As Long
    For r = 7 To hsLastRow
        Dim hsKP As String
        hsKP = Trim(CStr(hsWs.Cells(r, hsKPCol).Value))
        If hsKP = "" Then GoTo NextHSRow1
        
        If bhpDates.Exists(hsKP) Then
            Dim hsDate As Variant
            hsDate = hsWs.Cells(r, hsDateCol).Value
            If IsDate(hsDate) Then
                If CDate(hsDate) <> bhpDates(hsKP) Then
                    ' 日程変更あり → 起点を更新（より早い日付）
                    Dim changeDate As Date
                    changeDate = bhpDates(hsKP)
                    If changeDate < CDate(hsDate) Then changeDate = CDate(hsDate)
                    If changeDate < replaceFrom Then
                        replaceFrom = changeDate
                    End If
                End If
            End If
        End If
NextHSRow1:
    Next r
    
    ' 差し替え起点: replaceFrom
    
    ' 差し替え実行
    Dim replacedCount As Long
    Dim skippedR As Long
    Dim skippedN As Long
    Dim skippedSAY As Long
    replacedCount = 0
    skippedR = 0
    skippedN = 0
    skippedSAY = 0
    
    For r = 7 To hsLastRow
        ' 出荷日が起点より前ならスキップ
        Dim rowDate As Variant
        rowDate = hsWs.Cells(r, hsDateCol).Value
        If IsDate(rowDate) Then
            If CDate(rowDate) < replaceFrom Then GoTo NextHSRow2
        End If
        
        ' スキップ条件: メンテ列に「M」が入っている行
        ' V8: R列(18)=メンテ、V9: N列(14)=メンテ
        Dim menteCol As Long
        If modelType = "V8" Then
            menteCol = 18  ' V8のR列
        Else
            menteCol = 14  ' V9のN列
        End If
        Dim menteVal As String
        menteVal = Trim(CStr(hsWs.Cells(r, menteCol).Value))
        If UCase(menteVal) = "M" Then
            skippedN = skippedN + 1
            GoTo NextHSRow2
        End If
        
        ' スキップ条件: LAZ完了計画列に日付が入っている行
        ' V8: N列(14)=LAZ完了計画、V9: K列(11)=LAZ完了計画
        Dim lazCol As Long
        If modelType = "V8" Then
            lazCol = 14  ' V8のN列
        Else
            lazCol = 11  ' V9のK列
        End If
        Dim lazVal As Variant
        lazVal = hsWs.Cells(r, lazCol).Value
        If IsDate(lazVal) Then
            skippedR = skippedR + 1
            GoTo NextHSRow2
        End If
        
        ' スキップ条件3: S列～AY列に日付が入っている
        Dim hasDateInSAY As Boolean
        hasDateInSAY = False
        Dim c As Long
        For c = 19 To 51  ' S列(19)～AY列(51)
            If IsDate(hsWs.Cells(r, c).Value) Then
                hasDateInSAY = True
                Exit For
            End If
        Next c
        If hasDateInSAY Then
            skippedSAY = skippedSAY + 1
            ' S-AY列日付ありスキップ
            GoTo NextHSRow2
        End If
        
        ' 差し替え: BHPlanから対応するデータを転記
        Dim hsKP2 As String
        hsKP2 = Trim(CStr(hsWs.Cells(r, hsKPCol).Value))
        
        If hsKP2 <> "" Then
            ' BHPlanから該当行を探す
            Dim bhpRow As Long
            bhpRow = BHPlan行検索(targetWs, hsKP2, modelType, bhpLastRow)
            
            If bhpRow > 0 Then
                ' BHPlanのデータで星取表を更新
                If modelType = "V8" Then
                    ' V8: BH型式→I列, 順序確定日→J列, 出荷日→K列
                    hsWs.Cells(r, 9).Value = targetWs.Cells(bhpRow, g_ColBHType).Value
                    hsWs.Cells(r, 10).Value = targetWs.Cells(bhpRow, g_ColJunjoHakkoDate).Value
                    hsWs.Cells(r, 11).Value = targetWs.Cells(bhpRow, g_ColShukkaDate).Value
                Else
                    ' V9: BH型式→F列, 順序確定日→G列, 出荷日→H列
                    hsWs.Cells(r, 6).Value = targetWs.Cells(bhpRow, g_ColBHType).Value
                    hsWs.Cells(r, 7).Value = targetWs.Cells(bhpRow, g_ColJunjoHakkoDate).Value
                    hsWs.Cells(r, 8).Value = targetWs.Cells(bhpRow, g_ColShukkaDate).Value
                End If
                replacedCount = replacedCount + 1
            End If
        End If
NextHSRow2:
    Next r
    
    psWb.Save
    psWb.Close SaveChanges:=False
    
    Call ログ書込("Step19", "完了", _
        modelType & " 差替:" & replacedCount & "行, スキップ(LAZ日付):" & skippedR & _
        ", スキップ(メンテM):" & skippedN & ", スキップ(S-AY日付):" & skippedSAY)
    
    星取表差替処理 = replacedCount
End Function

' ============================================================
' BHPlanからKP-Noで行を検索
' ============================================================
Private Function BHPlan行検索(ws As Worksheet, kpNo As String, modelType As String, lastRow As Long) As Long
    BHPlan行検索 = 0
    Dim i As Long
    For i = g_DataStartRow To lastRow
        If Trim(CStr(ws.Cells(i, g_ColModel).Value)) = modelType Then
            If Trim(CStr(ws.Cells(i, g_ColKPNo).Value)) = kpNo Then
                BHPlan行検索 = i
                Exit Function
            End If
        End If
    Next i
End Function

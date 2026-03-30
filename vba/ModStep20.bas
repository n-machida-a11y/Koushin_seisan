Attribute VB_Name = "ModStep20"
Option Explicit

' ============================================================
' ステップ⑳: BH出荷・完了グラフ更新
'
' V8 Production Scheduleの「BH出荷・完了グラフ」シートを更新
' 月ごとのV8/V9出荷台数を集計して書き込む
' ============================================================
Public Sub Step20_グラフ更新(targetWs As Worksheet)
    If g_V8ProdSchedulePath = "" Then
        Call ログ書込("Step20_グラフ更新", "警告", "V8_ProductionScheduleパスが未設定です")
        Exit Sub
    End If
    
    ' BHPlanから月ごとの出荷台数を集計
    Dim v8Monthly As Object
    Set v8Monthly = CreateObject("Scripting.Dictionary")  ' "YYYY/MM" -> count
    Dim v9Monthly As Object
    Set v9Monthly = CreateObject("Scripting.Dictionary")
    
    Dim lastRow As Long
    lastRow = targetWs.Cells(targetWs.Rows.Count, 1).End(xlUp).Row
    
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
        
        Dim monthKey As String
        monthKey = Format(CDate(shukkaDate), "YYYY/MM")
        
        If model = "V8" Then
            If v8Monthly.Exists(monthKey) Then
                v8Monthly(monthKey) = v8Monthly(monthKey) + CLng(suryo)
            Else
                v8Monthly.Add monthKey, CLng(suryo)
            End If
        Else
            If v9Monthly.Exists(monthKey) Then
                v9Monthly(monthKey) = v9Monthly(monthKey) + CLng(suryo)
            Else
                v9Monthly.Add monthKey, CLng(suryo)
            End If
        End If
NextRow:
    Next i
    
    ' グラフシートを更新
    Dim psWb As Workbook
    Set psWb = Workbooks.Open(g_V8ProdSchedulePath)
    
    Dim graphWs As Worksheet
    Set graphWs = シート検索(psWb, g_SheetBHGraph)
    If graphWs Is Nothing Then
        Call ログ書込("Step20_グラフ更新", "エラー", "BH出荷・完了グラフシートが見つかりません（設定: " & g_SheetBHGraph & "）")
        psWb.Close SaveChanges:=False
        Exit Sub
    End If
    
    ' グラフシートのA列=出荷月, B列=V8-BH, C列=V9-BH, D列=BH合計
    Dim graphLastRow As Long
    graphLastRow = graphWs.Cells(graphWs.Rows.Count, 1).End(xlUp).Row
    
    ' 既存の月 -> 行マップ
    Dim monthRowMap As Object
    Set monthRowMap = CreateObject("Scripting.Dictionary")
    Dim r As Long
    For r = 2 To graphLastRow
        Dim monthVal As Variant
        monthVal = graphWs.Cells(r, 1).Value
        If Not IsEmpty(monthVal) Then
            monthRowMap(CStr(monthVal)) = r
        End If
    Next r
    
    ' 当月+4ヶ月分を更新
    Dim writtenCount As Long
    writtenCount = 0
    Dim m As Long
    For m = 0 To 4
        Dim targetMonth As Date
        targetMonth = DateSerial(Year(g_BaseDate), Month(g_BaseDate) + m, 1)
        Dim mk As String
        mk = Format(targetMonth, "YYYY/MM")
        
        ' グラフシートの月フォーマットに合わせる（例: "26/03"）
        Dim graphMK As String
        graphMK = Format(targetMonth, "YY/MM")
        
        Dim graphRow As Long
        If monthRowMap.Exists(graphMK) Then
            graphRow = monthRowMap(graphMK)
        Else
            ' 新しい行を追加
            graphLastRow = graphLastRow + 1
            graphRow = graphLastRow
            graphWs.Cells(graphRow, 1).Value = graphMK
        End If
        
        Dim v8Count As Long
        Dim v9Count As Long
        v8Count = 0
        v9Count = 0
        If v8Monthly.Exists(mk) Then v8Count = v8Monthly(mk)
        If v9Monthly.Exists(mk) Then v9Count = v9Monthly(mk)
        
        graphWs.Cells(graphRow, 2).Value = v8Count    ' B列: V8-BH
        graphWs.Cells(graphRow, 3).Value = v9Count    ' C列: V9-BH
        graphWs.Cells(graphRow, 4).Value = v8Count + v9Count  ' D列: BH合計
        
        writtenCount = writtenCount + 1
    Next m
    
    ' --- 表2: 日次データ（Row28～）の更新 ---
    ' A=日付, B=V8出荷, C=V8完了, D=V9出荷, E=V9完了, F=BH出荷, G=BH完了
    ' 出荷・完了計画から当月+4ヶ月の日次データを取得して転記
    Dim table2Start As Long
    table2Start = 0
    Dim tr As Long
    For tr = 27 To 30
        Dim hVal As String
        hVal = Trim(CStr(graphWs.Cells(tr, 1).Value))
        If hVal = "日" Or InStr(hVal, "日") > 0 Then
            table2Start = tr + 1
            Exit For
        End If
        ' 日付が入っていれば表2開始
        If IsDate(graphWs.Cells(tr, 1).Value) Then
            table2Start = tr
            Exit For
        End If
    Next tr
    
    If table2Start > 0 Then
        ' 表2の既存日付→行マップ
        Dim t2RowMap As Object
        Set t2RowMap = CreateObject("Scripting.Dictionary")
        Dim t2Last As Long
        t2Last = graphWs.Cells(graphWs.Rows.Count, 1).End(xlUp).Row
        For tr = table2Start To t2Last
            Dim t2Date As Variant
            t2Date = graphWs.Cells(tr, 1).Value
            If IsDate(t2Date) Then
                t2RowMap(Format(CDate(t2Date), "YYYY/MM/DD")) = tr
            End If
        Next tr
        
        ' BHPlanの日付ごとのV8/V9出荷台数（既に集計済みのv8Monthly/v9Monthlyは月次）
        ' 日次データが必要なので再集計
        Dim v8Daily As Object
        Set v8Daily = CreateObject("Scripting.Dictionary")
        Dim v9Daily As Object
        Set v9Daily = CreateObject("Scripting.Dictionary")
        
        Dim j As Long
        For j = g_DataStartRow To lastRow
            Dim mdl As String
            mdl = Trim(CStr(targetWs.Cells(j, g_ColModel).Value))
            If mdl <> "V8" And mdl <> "V9" Then GoTo NextDailyRow
            Dim sd As Variant
            sd = targetWs.Cells(j, g_ColShukkaDate).Value
            If IsEmpty(sd) Or Not IsDate(sd) Then GoTo NextDailyRow
            If CDate(sd) < g_BaseDate Then GoTo NextDailyRow
            Dim sq As Variant
            sq = targetWs.Cells(j, g_ColSuryo).Value
            If IsEmpty(sq) Or Not IsNumeric(sq) Then GoTo NextDailyRow
            Dim ddk As String
            ddk = Format(CDate(sd), "YYYY/MM/DD")
            If mdl = "V8" Then
                If v8Daily.Exists(ddk) Then
                    v8Daily(ddk) = v8Daily(ddk) + CLng(sq)
                Else
                    v8Daily.Add ddk, CLng(sq)
                End If
            Else
                If v9Daily.Exists(ddk) Then
                    v9Daily(ddk) = v9Daily(ddk) + CLng(sq)
                Else
                    v9Daily.Add ddk, CLng(sq)
                End If
            End If
NextDailyRow:
        Next j
        
        ' 表2に書き込み
        Dim dailyWritten As Long
        dailyWritten = 0
        Dim dKey As Variant
        For Each dKey In v8Daily.Keys
            If t2RowMap.Exists(CStr(dKey)) Then
                Dim dRow As Long
                dRow = t2RowMap(CStr(dKey))
                Dim v8d As Long
                Dim v9d As Long
                v8d = v8Daily(dKey)
                v9d = 0
                If v9Daily.Exists(CStr(dKey)) Then v9d = v9Daily(CStr(dKey))
                graphWs.Cells(dRow, 2).Value = v8d  ' B: V8出荷
                graphWs.Cells(dRow, 4).Value = v9d  ' D: V9出荷
                graphWs.Cells(dRow, 6).Value = v8d + v9d  ' F: BH出荷台数
                dailyWritten = dailyWritten + 1
            End If
        Next dKey
        ' V9のみの日付も処理
        For Each dKey In v9Daily.Keys
            If Not v8Daily.Exists(CStr(dKey)) Then
                If t2RowMap.Exists(CStr(dKey)) Then
                    dRow = t2RowMap(CStr(dKey))
                    graphWs.Cells(dRow, 4).Value = v9Daily(dKey)
                    graphWs.Cells(dRow, 6).Value = v9Daily(dKey)
                    dailyWritten = dailyWritten + 1
                End If
            End If
        Next dKey
    End If
    
    psWb.Save
    psWb.Close SaveChanges:=False
    
    Call ログ書込("Step20_グラフ更新", "完了", writtenCount & "ヶ月分のグラフデータを更新しました")
End Sub

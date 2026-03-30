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
    On Error Resume Next
    Set graphWs = psWb.Sheets(g_SheetBHGraph)
    On Error GoTo 0
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
    
    psWb.Save
    psWb.Close SaveChanges:=False
    
    Call ログ書込("Step20_グラフ更新", "完了", writtenCount & "ヶ月分のグラフデータを更新しました")
End Sub

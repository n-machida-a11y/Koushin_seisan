Attribute VB_Name = "ModStep25"
Option Explicit

' ============================================================
' Step25: KMP MPS更新
'
' KMP MPSシートの末尾に新しいブロックを追加する。
' 構成: タイトル行 → V8セクション(ヘッダー+データ+TOTAL) → V9セクション(同)
' 前回ブロックとの変更点はAC列(29列目)に記録する。
'
' V8ヘッダー: LAA101,LAF001,LAF002,LAF003,LAG001,LAG002,LAE101,
'             LAE102,LAE201,LBL001,LBL002,LAL001,LAN001,LAN002,
'             LCN001,LAC001,LAC002,LCC001,LCC002,LAD001,LAD002,
'             LCD001,LCD002,TOTAL
' V9ヘッダー: LAA001,LAF001,LAG001,LAE001,LAL001,LAN001,LAC001,LAD001,TOTAL
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
        Call ログ書込("Step25", "エラー", "KMP MPSシートが見つかりません")
        psWb.Close SaveChanges:=False
        Exit Sub
    End If
    
    ' --- 前回ブロックの位置を特定（末尾から逆走して最後のV8 TOTALを探す） ---
    Dim prevV8Start As Long
    Dim prevV9Start As Long
    prevV8Start = 0
    prevV9Start = 0
    
    Dim mpsLastRow As Long
    mpsLastRow = mpsWs.Cells(mpsWs.Rows.Count, 2).End(xlUp).Row
    
    ' 最後のV9 TOTALから遡ってV9開始、V8開始を見つける
    Dim r As Long
    For r = mpsLastRow To 1 Step -1
        Dim bVal As String
        bVal = Trim(CStr(mpsWs.Cells(r, 2).Value))
        If bVal = "V9" And prevV9Start = 0 Then
            prevV9Start = r
        End If
        If bVal = "V8" And prevV8Start = 0 Then
            prevV8Start = r
            Exit For
        End If
    Next r
    
    ' --- 前回V8データを読み取り（比較用） ---
    Dim prevV8Data As Object
    Set prevV8Data = CreateObject("Scripting.Dictionary")
    If prevV8Start > 0 Then
        Dim pdr As Long
        For pdr = prevV8Start + 2 To mpsLastRow
            If Trim(CStr(mpsWs.Cells(pdr, 2).Value)) = "TOTAL" Then Exit For
            If Trim(CStr(mpsWs.Cells(pdr, 2).Value)) = "V9" Then Exit For
            Dim pdDate As Variant
            pdDate = mpsWs.Cells(pdr, 3).Value
            If IsDate(pdDate) Then
                Dim pdKey As String
                pdKey = Format(CDate(pdDate), "YYYY/MM/DD")
                prevV8Data(pdKey) = mpsWs.Cells(pdr, 27).Value  ' TOTAL列(27)
            End If
        Next pdr
    End If
    
    ' --- BHPlanからKMP出荷日ごとのコンポーネント台数を集計 ---
    ' BHPlanのBP列以降(LAA101Q等)をshipment date(N列)で集計
    Dim bhpLastRow As Long
    bhpLastRow = targetWs.Cells(targetWs.Rows.Count, 1).End(xlUp).Row
    
    ' V8コンポーネントヘッダー（0110BHPlanの日程表のBP列以降に対応）
    Dim v8Comps As Variant
    v8Comps = Array("LAA101", "LAF001", "LAF002", "LAF003", "LAG001", "LAG002", _
                    "LAE101", "LAE102", "LAE201", "LBL001", "LBL002", "LAL001", _
                    "LAN001", "LAN002", "LCN001", "LAC001", "LAC002", "LCC001", _
                    "LCC002", "LAD001", "LAD002", "LCD001", "LCD002")
    
    Dim v9Comps As Variant
    v9Comps = Array("LAA001", "LAF001", "LAG001", "LAE001", "LAL001", "LAN001", _
                    "LAC001", "LAD001")
    
    ' BHPlanの日付ごとV8/V9台数集計
    Dim v8Dates As Object
    Set v8Dates = CreateObject("Scripting.Dictionary")
    Dim v9Dates As Object
    Set v9Dates = CreateObject("Scripting.Dictionary")
    
    Dim i As Long
    For i = g_DataStartRow To bhpLastRow
        Dim model As String
        model = Trim(CStr(targetWs.Cells(i, g_ColModel).Value))
        If model <> "V8" And model <> "V9" Then GoTo NextBHP
        
        Dim shukkaDate As Variant
        shukkaDate = targetWs.Cells(i, g_ColShukkaDate).Value
        If IsEmpty(shukkaDate) Or Not IsDate(shukkaDate) Then GoTo NextBHP
        If CDate(shukkaDate) < g_BaseDate Then GoTo NextBHP
        
        Dim suryo As Variant
        suryo = targetWs.Cells(i, g_ColSuryo).Value
        If IsEmpty(suryo) Or Not IsNumeric(suryo) Then GoTo NextBHP
        
        Dim dk As String
        dk = Format(CDate(shukkaDate), "YYYY/MM/DD")
        
        If model = "V8" Then
            If Not v8Dates.Exists(dk) Then v8Dates.Add dk, CLng(suryo) Else v8Dates(dk) = v8Dates(dk) + CLng(suryo)
        Else
            If Not v9Dates.Exists(dk) Then v9Dates.Add dk, CLng(suryo) Else v9Dates(dk) = v9Dates(dk) + CLng(suryo)
        End If
NextBHP:
    Next i
    
    ' --- 新しいブロックを末尾に追加 ---
    Dim writeRow As Long
    writeRow = mpsLastRow + 3  ' 空行2行あけてから
    
    ' タイトル行
    mpsWs.Cells(writeRow, 2).Value = " KMP NEW MPS（" & Format(Date, "M/DD/YYYY") & "）"
    writeRow = writeRow + 2
    
    ' --- V8セクション ---
    mpsWs.Cells(writeRow, 2).Value = "V8"
    ' 部品番号行はスキップ（前回と同じなので）
    writeRow = writeRow + 1
    
    ' ヘッダー行
    mpsWs.Cells(writeRow, 2).Value = "MPS date"
    mpsWs.Cells(writeRow, 3).Value = "KMP Shipment"
    Dim ci As Long
    For ci = 0 To UBound(v8Comps)
        mpsWs.Cells(writeRow, 4 + ci).Value = v8Comps(ci)
    Next ci
    mpsWs.Cells(writeRow, 4 + UBound(v8Comps) + 1).Value = "TOTAL"
    Dim v8TotalCol As Long
    v8TotalCol = 4 + UBound(v8Comps) + 1  ' TOTAL列
    writeRow = writeRow + 1
    
    ' データ行（日付順）
    Dim sortedDates() As String
    Dim dateKeys As Variant
    dateKeys = v8Dates.Keys
    ' 簡易ソート
    Dim sd1 As Long
    Dim sd2 As Long
    Dim tmpStr As String
    For sd1 = 0 To v8Dates.Count - 2
        For sd2 = sd1 + 1 To v8Dates.Count - 1
            If dateKeys(sd1) > dateKeys(sd2) Then
                tmpStr = dateKeys(sd1)
                dateKeys(sd1) = dateKeys(sd2)
                dateKeys(sd2) = tmpStr
            End If
        Next sd2
    Next sd1
    
    Dim v8TotalRow As Long
    Dim totalSums() As Long
    ReDim totalSums(0 To UBound(v8Comps))
    
    For sd1 = 0 To v8Dates.Count - 1
        dk = dateKeys(sd1)
        mpsWs.Cells(writeRow, 3).Value = CDate(dk)
        ' 各コンポーネントには台数を均等に配分（簡易版: TOTALのみ設定）
        mpsWs.Cells(writeRow, v8TotalCol).Value = v8Dates(dk)
        
        ' 前回との比較
        If prevV8Data.Exists(dk) Then
            Dim diff As Long
            diff = CLng(v8Dates(dk)) - CLng(prevV8Data(dk))
            If diff <> 0 Then
                mpsWs.Cells(writeRow, 28).Value = "Change  qity->"
                mpsWs.Cells(writeRow, 29).Value = diff
            End If
        Else
            mpsWs.Cells(writeRow, 28).Value = "Add ship date->"
            mpsWs.Cells(writeRow, 29).Value = v8Dates(dk)
        End If
        
        writeRow = writeRow + 1
    Next sd1
    
    ' TOTAL行
    mpsWs.Cells(writeRow, 2).Value = "TOTAL"
    ' TOTAL列に合計
    Dim v8Sum As Long
    v8Sum = 0
    Dim dkv As Variant
    For Each dkv In v8Dates.Keys
        v8Sum = v8Sum + v8Dates(dkv)
    Next dkv
    mpsWs.Cells(writeRow, v8TotalCol).Value = v8Sum
    writeRow = writeRow + 3
    
    ' --- V9セクション ---
    mpsWs.Cells(writeRow, 2).Value = "V9"
    writeRow = writeRow + 1
    
    ' ヘッダー行
    mpsWs.Cells(writeRow, 2).Value = "MPS date"
    mpsWs.Cells(writeRow, 3).Value = "KMP Shipment"
    For ci = 0 To UBound(v9Comps)
        mpsWs.Cells(writeRow, 4 + ci).Value = v9Comps(ci)
    Next ci
    Dim v9TotalCol As Long
    v9TotalCol = 4 + UBound(v9Comps) + 1
    mpsWs.Cells(writeRow, v9TotalCol).Value = "TOTAL"
    writeRow = writeRow + 1
    
    ' V9データ行
    dateKeys = v9Dates.Keys
    For sd1 = 0 To v9Dates.Count - 2
        For sd2 = sd1 + 1 To v9Dates.Count - 1
            If dateKeys(sd1) > dateKeys(sd2) Then
                tmpStr = dateKeys(sd1)
                dateKeys(sd1) = dateKeys(sd2)
                dateKeys(sd2) = tmpStr
            End If
        Next sd2
    Next sd1
    
    For sd1 = 0 To v9Dates.Count - 1
        dk = dateKeys(sd1)
        mpsWs.Cells(writeRow, 3).Value = CDate(dk)
        mpsWs.Cells(writeRow, v9TotalCol).Value = v9Dates(dk)
        writeRow = writeRow + 1
    Next sd1
    
    ' V9 TOTAL行
    mpsWs.Cells(writeRow, 2).Value = "TOTAL"
    Dim v9Sum As Long
    v9Sum = 0
    For Each dkv In v9Dates.Keys
        v9Sum = v9Sum + v9Dates(dkv)
    Next dkv
    mpsWs.Cells(writeRow, v9TotalCol).Value = v9Sum
    
    psWb.Save
    psWb.Close SaveChanges:=False
    
    Call ログ書込("Step25_KMP_MPS更新", "完了", _
        "KMP MPSに新ブロック追加（V8: " & v8Dates.Count & "日分、V9: " & v9Dates.Count & "日分）")
End Sub

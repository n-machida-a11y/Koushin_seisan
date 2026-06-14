Attribute VB_Name = "ModStep25"
Option Explicit

' ============================================================
' Step25: KMP MPS更新
'
' KMP MPSシートの末尾に新しいブロックを追加する。
' V8: D～Z列にコンポーネント、AA列(27)にTOTAL(SUM数式)
' V9: D,E,H,J,O,P,S,W列にコンポーネント、AA列(27)にTOTAL(SUM数式)
' AC列(29): 前回ブロックとの差分
' フォント: MSPゴシック、罫線: ヘッダーthin/データhair/TOTALthin
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
    
    ' --- 前回V8ブロックの位置特定 ---
    Dim mpsLastRow As Long
    mpsLastRow = mpsWs.Cells(mpsWs.Rows.Count, 2).End(xlUp).Row
    
    Dim prevV8DataStart As Long
    prevV8DataStart = 0
    Dim prevV8DataEnd As Long
    prevV8DataEnd = 0
    
    ' 末尾から逆走してV8セクションを見つける
    Dim r As Long
    For r = mpsLastRow To 1 Step -1
        Dim bv As String
        bv = Trim(CStr(mpsWs.Cells(r, 2).Value))
        If bv = "V8" Then
            ' V8ヘッダーの2行下からデータ
            prevV8DataStart = r + 2
            Exit For
        End If
    Next r
    ' V8 TOTAL行を探す
    If prevV8DataStart > 0 Then
        For r = prevV8DataStart To mpsLastRow
            If Trim(CStr(mpsWs.Cells(r, 2).Value)) = "TOTAL" Then
                prevV8DataEnd = r - 1
                Exit For
            End If
            If Trim(CStr(mpsWs.Cells(r, 2).Value)) = "V9" Then
                prevV8DataEnd = r - 1
                Exit For
            End If
        Next r
    End If
    
    ' 前回V8の出荷日→TOTAL値マップ
    Dim prevV8Map As Object
    Set prevV8Map = CreateObject("Scripting.Dictionary")
    If prevV8DataStart > 0 And prevV8DataEnd > 0 Then
        For r = prevV8DataStart To prevV8DataEnd
            Dim pdDate As Variant
            pdDate = mpsWs.Cells(r, 3).Value
            If IsDate(pdDate) Then
                prevV8Map(Format(CDate(pdDate), "YYYY/MM/DD")) = r
            End If
        Next r
    End If
    
    ' --- BHPlanからV8/V9の出荷日別台数を集計 ---
    Dim bhpLastRow As Long
    bhpLastRow = targetWs.Cells(targetWs.Rows.Count, 1).End(xlUp).Row
    
    Dim v8Dates As Object
    Set v8Dates = CreateObject("Scripting.Dictionary")
    Dim v9Dates As Object
    Set v9Dates = CreateObject("Scripting.Dictionary")
    
    Dim i As Long
    For i = g_DataStartRow To bhpLastRow
        Dim model As String
        model = Trim(CStr(targetWs.Cells(i, g_ColModel).Value))
        If model <> "V8" And model <> "V9" Then GoTo NextBHP
        Dim sd As Variant
        sd = targetWs.Cells(i, g_ColShukkaDate).Value
        If IsEmpty(sd) Or Not IsDate(sd) Then GoTo NextBHP
        If CDate(sd) < g_BaseDate Then GoTo NextBHP
        Dim suryo As Variant
        suryo = targetWs.Cells(i, g_ColSuryo).Value
        If IsEmpty(suryo) Or Not IsNumeric(suryo) Then GoTo NextBHP
        Dim dk As String
        dk = Format(CDate(sd), "YYYY/MM/DD")
        If model = "V8" Then
            If Not v8Dates.Exists(dk) Then v8Dates.Add dk, CLng(suryo) Else v8Dates(dk) = v8Dates(dk) + CLng(suryo)
        Else
            If Not v9Dates.Exists(dk) Then v9Dates.Add dk, CLng(suryo) Else v9Dates(dk) = v9Dates(dk) + CLng(suryo)
        End If
NextBHP:
    Next i
    
    ' 日付をソート
    Dim v8Keys As Variant
    v8Keys = v8Dates.Keys
    Dim s1 As Long
    Dim s2 As Long
    Dim tmp As String
    For s1 = 0 To v8Dates.Count - 2
        For s2 = s1 + 1 To v8Dates.Count - 1
            If v8Keys(s1) > v8Keys(s2) Then
                tmp = v8Keys(s1): v8Keys(s1) = v8Keys(s2): v8Keys(s2) = tmp
            End If
        Next s2
    Next s1
    Dim v9Keys As Variant
    v9Keys = v9Dates.Keys
    For s1 = 0 To v9Dates.Count - 2
        For s2 = s1 + 1 To v9Dates.Count - 1
            If v9Keys(s1) > v9Keys(s2) Then
                tmp = v9Keys(s1): v9Keys(s1) = v9Keys(s2): v9Keys(s2) = tmp
            End If
        Next s2
    Next s1
    
    ' --- 新ブロック書き込み ---
    Dim wr As Long
    wr = mpsLastRow + 3
    
    ' タイトル行
    mpsWs.Cells(wr, 2).Value = " KMP NEW MPS（" & Format(Date, "M/DD/YYYY") & "）"
    mpsWs.Cells(wr, 2).Font.Name = "ＭＳ Ｐゴシック"
    mpsWs.Cells(wr, 2).Font.Size = 24
    mpsWs.Cells(wr, 2).Font.Bold = True
    wr = wr + 2
    
    ' ===== V8セクション =====
    ' V8ラベル
    mpsWs.Cells(wr, 2).Value = "V8"
    mpsWs.Cells(wr, 2).Font.Name = "ＭＳ Ｐゴシック"
    mpsWs.Cells(wr, 2).Font.Size = 22
    mpsWs.Cells(wr, 2).Font.Bold = True
    ' 比較ヘッダー
    mpsWs.Cells(wr, 29).Value = "Comparison with the previous"
    mpsWs.Cells(wr, 29).Font.Name = "ＭＳ Ｐゴシック"
    mpsWs.Cells(wr, 29).Font.Size = 22
    wr = wr + 1
    
    ' V8ヘッダー行
    Dim v8Headers As Variant
    v8Headers = Array("MPS date", "KMP " & vbLf & "Shipment", _
        "LAA101", "LAF001", "LAF002", "LAF003", "LAG001", "LAG002", _
        "LAE101", "LAE102", "LAE201", "LBL001", "LBL002", "LAL001", _
        "LAN001", "LAN002", "LCN001", "LAC001", "LAC002", "LCC001", _
        "LCC002", "LAD001", "LAD002", "LCD001", "LCD002")
    
    Dim c As Long
    For c = 0 To UBound(v8Headers)
        Dim hCell As Range
        Set hCell = mpsWs.Cells(wr, 2 + c)
        hCell.Value = v8Headers(c)
        hCell.Font.Name = "ＭＳ Ｐゴシック"
        If c <= 1 Then
            hCell.Font.Size = 12
        Else
            hCell.Font.Size = 19
        End If
        hCell.Borders(xlEdgeBottom).LineStyle = xlContinuous
        hCell.Borders(xlEdgeBottom).Weight = xlThin
        hCell.Borders(xlEdgeLeft).LineStyle = xlContinuous
        hCell.Borders(xlEdgeLeft).Weight = xlThin
    Next c
    ' TOTAL列(AA=27)
    mpsWs.Cells(wr, 27).Value = "TOTAL"
    mpsWs.Cells(wr, 27).Font.Name = "ＭＳ Ｐゴシック"
    mpsWs.Cells(wr, 27).Font.Size = 16
    mpsWs.Cells(wr, 27).Interior.Color = RGB(255, 255, 0)
    mpsWs.Cells(wr, 27).Borders(xlEdgeBottom).LineStyle = xlContinuous
    mpsWs.Cells(wr, 27).Borders(xlEdgeBottom).Weight = xlThin
    ' 比較列ヘッダー
    mpsWs.Cells(wr, 29).Value = v8Headers(2)  ' LAA101
    mpsWs.Cells(wr, 29).Font.Name = "ＭＳ Ｐゴシック"
    mpsWs.Cells(wr, 29).Font.Size = 19
    mpsWs.Cells(wr, 29).Borders(xlEdgeBottom).LineStyle = xlContinuous
    mpsWs.Cells(wr, 29).Borders(xlEdgeBottom).Weight = xlThin
    wr = wr + 1
    
    ' V8データ行
    Dim v8DataStart As Long
    v8DataStart = wr
    For s1 = 0 To v8Dates.Count - 1
        dk = v8Keys(s1)
        mpsWs.Cells(wr, 3).Value = CDate(dk)
        mpsWs.Cells(wr, 3).NumberFormat = "YYYY/M/D"
        ' TOTAL列にSUM数式
        mpsWs.Cells(wr, 27).Formula = "=SUM(D" & wr & ":Z" & wr & ")"
        mpsWs.Cells(wr, 27).Interior.Color = RGB(255, 255, 0)
        
        ' スタイル設定
        For c = 2 To 27
            mpsWs.Cells(wr, c).Font.Name = "ＭＳ Ｐゴシック"
            If c = 2 Then
                mpsWs.Cells(wr, c).Font.Size = 12
            Else
                mpsWs.Cells(wr, c).Font.Size = 18
            End If
            mpsWs.Cells(wr, c).Borders(xlEdgeBottom).LineStyle = xlContinuous
            mpsWs.Cells(wr, c).Borders(xlEdgeBottom).Weight = xlHairline
            mpsWs.Cells(wr, c).Borders(xlEdgeLeft).LineStyle = xlContinuous
            mpsWs.Cells(wr, c).Borders(xlEdgeLeft).Weight = xlThin
        Next c
        
        ' 前回との差分
        If prevV8Map.Exists(dk) Then
            Dim prevRow As Long
            prevRow = prevV8Map(dk)
            mpsWs.Cells(wr, 29).Formula = "=D" & wr & "-D" & prevRow
            mpsWs.Cells(wr, 28).Value = "Change  qity->"
        Else
            mpsWs.Cells(wr, 29).Value = v8Dates(dk)
            mpsWs.Cells(wr, 28).Value = "Add ship date->"
        End If
        mpsWs.Cells(wr, 28).Font.Name = "ＭＳ Ｐゴシック"
        mpsWs.Cells(wr, 28).Font.Size = 14
        mpsWs.Cells(wr, 28).Font.Bold = True
        mpsWs.Cells(wr, 29).Font.Name = "ＭＳ Ｐゴシック"
        mpsWs.Cells(wr, 29).Font.Size = 18
        mpsWs.Cells(wr, 29).Borders(xlEdgeBottom).LineStyle = xlContinuous
        mpsWs.Cells(wr, 29).Borders(xlEdgeBottom).Weight = xlHairline
        mpsWs.Cells(wr, 29).Borders(xlEdgeLeft).LineStyle = xlContinuous
        mpsWs.Cells(wr, 29).Borders(xlEdgeLeft).Weight = xlThin
        
        wr = wr + 1
    Next s1
    Dim v8DataEnd As Long
    v8DataEnd = wr - 1
    
    ' V8 TOTAL行
    mpsWs.Cells(wr, 2).Value = "TOTAL"
    mpsWs.Cells(wr, 2).Font.Name = "ＭＳ Ｐゴシック"
    mpsWs.Cells(wr, 2).Font.Size = 16
    ' D～Z列にSUM数式
    For c = 4 To 26
        mpsWs.Cells(wr, c).Formula = "=SUM(" & Chr(64 + c) & v8DataStart & ":" & Chr(64 + c) & v8DataEnd & ")"
        mpsWs.Cells(wr, c).Font.Name = "ＭＳ Ｐゴシック"
        mpsWs.Cells(wr, c).Font.Size = 18
        mpsWs.Cells(wr, c).Borders(xlEdgeBottom).LineStyle = xlContinuous
        mpsWs.Cells(wr, c).Borders(xlEdgeBottom).Weight = xlThin
        mpsWs.Cells(wr, c).Borders(xlEdgeLeft).LineStyle = xlContinuous
        mpsWs.Cells(wr, c).Borders(xlEdgeLeft).Weight = xlThin
    Next c
    ' AA列TOTAL
    mpsWs.Cells(wr, 27).Formula = "=SUM(AA" & v8DataStart & ":AA" & v8DataEnd & ")"
    mpsWs.Cells(wr, 27).Font.Name = "ＭＳ Ｐゴシック"
    mpsWs.Cells(wr, 27).Font.Size = 18
    mpsWs.Cells(wr, 27).Borders(xlEdgeBottom).LineStyle = xlContinuous
    mpsWs.Cells(wr, 27).Borders(xlEdgeBottom).Weight = xlThin
    ' AC列合計
    mpsWs.Cells(wr, 29).Formula = "=SUM(AC" & v8DataStart & ":AC" & v8DataEnd & ")"
    mpsWs.Cells(wr, 29).Font.Name = "ＭＳ Ｐゴシック"
    mpsWs.Cells(wr, 29).Font.Size = 18
    mpsWs.Cells(wr, 29).Interior.Color = RGB(255, 255, 0)
    wr = wr + 3
    
    ' ===== V9セクション =====
    mpsWs.Cells(wr, 2).Value = "V9"
    mpsWs.Cells(wr, 2).Font.Name = "ＭＳ Ｐゴシック"
    mpsWs.Cells(wr, 2).Font.Size = 20
    mpsWs.Cells(wr, 2).Font.Bold = True
    wr = wr + 1
    
    ' V9ヘッダー（V8と同じ列位置にマッピング: D,E,H,J,O,P,S,W列）
    Dim v9ColMap As Variant
    v9ColMap = Array(4, 5, 8, 10, 15, 16, 19, 23)  ' V9コンポーネントの列位置
    Dim v9Names As Variant
    v9Names = Array("LAA001", "LAF001", "LAG001", "LAE001", "LAL001", "LAN001", "LAC001", "LAD001")
    
    mpsWs.Cells(wr, 2).Value = "MPS date"
    mpsWs.Cells(wr, 2).Font.Name = "ＭＳ Ｐゴシック"
    mpsWs.Cells(wr, 2).Font.Size = 12
    mpsWs.Cells(wr, 2).Borders(xlEdgeBottom).LineStyle = xlContinuous
    mpsWs.Cells(wr, 2).Borders(xlEdgeBottom).Weight = xlThin
    mpsWs.Cells(wr, 3).Value = "KMP" & vbLf & "Shipment"
    mpsWs.Cells(wr, 3).Font.Name = "ＭＳ Ｐゴシック"
    mpsWs.Cells(wr, 3).Font.Size = 12
    mpsWs.Cells(wr, 3).Borders(xlEdgeBottom).LineStyle = xlContinuous
    mpsWs.Cells(wr, 3).Borders(xlEdgeBottom).Weight = xlThin
    
    Dim vi As Long
    For vi = 0 To UBound(v9ColMap)
        Dim vc As Long
        vc = v9ColMap(vi)
        mpsWs.Cells(wr, vc).Value = v9Names(vi)
        mpsWs.Cells(wr, vc).Font.Name = "ＭＳ Ｐゴシック"
        mpsWs.Cells(wr, vc).Font.Size = 19
        mpsWs.Cells(wr, vc).Borders(xlEdgeBottom).LineStyle = xlContinuous
        mpsWs.Cells(wr, vc).Borders(xlEdgeBottom).Weight = xlThin
    Next vi
    mpsWs.Cells(wr, 27).Value = "TOTAL"
    mpsWs.Cells(wr, 27).Font.Name = "ＭＳ Ｐゴシック"
    mpsWs.Cells(wr, 27).Font.Size = 14
    mpsWs.Cells(wr, 27).Borders(xlEdgeBottom).LineStyle = xlContinuous
    mpsWs.Cells(wr, 27).Borders(xlEdgeBottom).Weight = xlThin
    wr = wr + 1
    
    ' V9データ行
    Dim v9DataStart As Long
    v9DataStart = wr
    For s1 = 0 To v9Dates.Count - 1
        dk = v9Keys(s1)
        mpsWs.Cells(wr, 3).Value = CDate(dk)
        mpsWs.Cells(wr, 3).NumberFormat = "YYYY/M/D"
        mpsWs.Cells(wr, 27).Formula = "=SUM(D" & wr & ":Z" & wr & ")"
        
        ' スタイル
        mpsWs.Cells(wr, 2).Font.Name = "ＭＳ Ｐゴシック"
        mpsWs.Cells(wr, 2).Font.Size = 12
        mpsWs.Cells(wr, 3).Font.Name = "ＭＳ Ｐゴシック"
        mpsWs.Cells(wr, 3).Font.Size = 18
        For vi = 0 To UBound(v9ColMap)
            vc = v9ColMap(vi)
            mpsWs.Cells(wr, vc).Font.Name = "ＭＳ Ｐゴシック"
            mpsWs.Cells(wr, vc).Font.Size = 18
            mpsWs.Cells(wr, vc).Borders(xlEdgeBottom).LineStyle = xlContinuous
            mpsWs.Cells(wr, vc).Borders(xlEdgeBottom).Weight = xlHairline
        Next vi
        mpsWs.Cells(wr, 27).Font.Name = "ＭＳ Ｐゴシック"
        mpsWs.Cells(wr, 27).Font.Size = 16
        mpsWs.Cells(wr, 2).Borders(xlEdgeBottom).LineStyle = xlContinuous
        mpsWs.Cells(wr, 2).Borders(xlEdgeBottom).Weight = xlHairline
        mpsWs.Cells(wr, 3).Borders(xlEdgeBottom).LineStyle = xlContinuous
        mpsWs.Cells(wr, 3).Borders(xlEdgeBottom).Weight = xlHairline
        mpsWs.Cells(wr, 27).Borders(xlEdgeBottom).LineStyle = xlContinuous
        mpsWs.Cells(wr, 27).Borders(xlEdgeBottom).Weight = xlHairline
        
        wr = wr + 1
    Next s1
    Dim v9DataEnd As Long
    v9DataEnd = wr - 1
    
    ' V9 TOTAL行
    mpsWs.Cells(wr, 2).Value = "TOTAL"
    mpsWs.Cells(wr, 2).Font.Name = "ＭＳ Ｐゴシック"
    mpsWs.Cells(wr, 2).Font.Size = 16
    For vi = 0 To UBound(v9ColMap)
        vc = v9ColMap(vi)
        Dim colLetter As String
        If vc <= 26 Then
            colLetter = Chr(64 + vc)
        Else
            colLetter = "A" & Chr(64 + vc - 26)
        End If
        mpsWs.Cells(wr, vc).Formula = "=SUM(" & colLetter & v9DataStart & ":" & colLetter & v9DataEnd & ")"
        mpsWs.Cells(wr, vc).Font.Name = "ＭＳ Ｐゴシック"
        mpsWs.Cells(wr, vc).Font.Size = 16
        mpsWs.Cells(wr, vc).Borders(xlEdgeBottom).LineStyle = xlContinuous
        mpsWs.Cells(wr, vc).Borders(xlEdgeBottom).Weight = xlThin
    Next vi
    mpsWs.Cells(wr, 27).Formula = "=SUM(AA" & v9DataStart & ":AA" & v9DataEnd & ")"
    mpsWs.Cells(wr, 27).Font.Name = "ＭＳ Ｐゴシック"
    mpsWs.Cells(wr, 27).Font.Size = 16
    mpsWs.Cells(wr, 27).Borders(xlEdgeBottom).LineStyle = xlContinuous
    mpsWs.Cells(wr, 27).Borders(xlEdgeBottom).Weight = xlThin
    
    Call 安全保存(psWb)
    psWb.Close SaveChanges:=False
    
    Call ログ書込("Step25_KMP_MPS更新", "完了", _
        "KMP MPSに新ブロック追加（V8: " & v8Dates.Count & "日分、V9: " & v9Dates.Count & "日分）")
End Sub

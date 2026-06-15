Attribute VB_Name = "ModStep13"
Option Explicit

' ============================================================
' ステップ⑬: V8/V9集計表作成
' (2026-06 正解Excel「4月の結果」に合わせて全面改修)
'
' BHPlan日程表からMODEL=V8/V9の行(出荷日が当月以降)を抽出し、
' V8 Production Scheduleの集計シートの日付ブロックを更新する。
'
' 集計シートの構造(0412集計で確認):
'   上部       : 年度別・月別ブロック(SUMIFSで日付ブロックを参照) → 触らない
'   日付ブロック: ヘッダー行 / 日付行(B列=日付) / 総計行(B列="総計")
'   列構成は月により変わるため、ヘッダー行から動的に列を解決する:
'     ・型式別詳細列 : ヘッダーが型式名と完全一致(VU-LD091, YB-LA001等)
'     ・グループ列   : ヘッダー1行目がLAZ/LAY型式と一致(LAZ201等)
'     ・合計/累計列  : 数式のため対象外
'
' 書込ルール:
'   ・V8行: S列(BH型式TYPE)/Y列(LAZ型式)/Z列(LAY型式)の3キーで列解決
'   ・V9行: S列(BH型式TYPE)で列解決(YB-LA001等はグループ列と詳細列の両方に一致)
'   ・数式セルは一切触らない(グループ列がSUMIFSで詳細列から自動集計される行がある)
'   ・過去月の値は保持し、当月以降のみクリア→書込(月次ローリング)
' 行追加:
'   ・新しい日付は日付順の位置に挿入(直上の行をコピーして数式・書式を継承)
'   ・範囲内への挿入なので総計SUM/年度別SUMIFSの参照範囲が自動拡張される
'   ・既存最終日より後の日付のみ総計行の直上に挿入し、総計行のSUM範囲を修復
' 罫線: 月の変わり目=二重線 / それ以外=点線(hair) で全日付行を引き直す
' ============================================================
Public Sub Step13_集計表作成(targetWs As Worksheet)
    If g_V8ProdSchedulePath = "" Then
        Call ログ書込("Step13_集計表作成", "警告", "V8_ProductionScheduleパスが未設定です")
        Exit Sub
    End If

    Dim psWb As Workbook
    Set psWb = Workbooks.Open(g_V8ProdSchedulePath)

    Dim shukeiWs As Worksheet
    Set shukeiWs = シート検索(psWb, g_SheetV8Shukei)
    ' フォールバック: シート名末尾が"集計"のシートを動的検索
    If shukeiWs Is Nothing Then
        Set shukeiWs = 集計シート動的検索(psWb)
        If Not shukeiWs Is Nothing Then
            Call ログ書込("Step13_集計表作成", "情報", _
                "集計シートを動的検索で発見: " & shukeiWs.Name)
        End If
    End If
    If shukeiWs Is Nothing Then
        Call ログ書込("Step13_集計表作成", "エラー", _
            "集計シートが見つかりません（設定: " & g_SheetV8Shukei & "、末尾""集計""のシートも無し）")
        psWb.Close SaveChanges:=False
        Exit Sub
    End If

    ' --- 日付ブロックの特定 ---
    Dim r As Long
    Dim lastB As Long
    lastB = shukeiWs.Cells(shukeiWs.Rows.Count, 2).End(xlUp).Row

    Dim totalRow As Long
    totalRow = 0
    For r = lastB To 2 Step -1
        If InStr(Trim(CStr(shukeiWs.Cells(r, 2).Value)), "総計") > 0 Then
            totalRow = r
            Exit For
        End If
    Next r
    If totalRow = 0 Then
        Call ログ書込("Step13_集計表作成", "エラー", "集計シートに総計行(B列)が見つかりません")
        psWb.Close SaveChanges:=False
        Exit Sub
    End If

    ' 総計行の直上から上へ連続する日付行が日付ブロック
    Dim firstDataRow As Long
    Dim lastDataRow As Long
    firstDataRow = 0
    lastDataRow = 0
    For r = totalRow - 1 To 2 Step -1
        If IsDate(shukeiWs.Cells(r, 2).Value) Then
            If lastDataRow = 0 Then lastDataRow = r
            firstDataRow = r
        ElseIf lastDataRow > 0 Then
            Exit For
        End If
    Next r
    If lastDataRow = 0 Then
        Call ログ書込("Step13_集計表作成", "エラー", "集計シートに日付行が見つかりません")
        psWb.Close SaveChanges:=False
        Exit Sub
    End If

    Dim headerRow As Long
    headerRow = firstDataRow - 1

    ' --- ヘッダーの読み取り(正規化) ---
    Dim lastHdrCol As Long
    lastHdrCol = shukeiWs.Cells(headerRow, shukeiWs.Columns.Count).End(xlToLeft).Column
    Dim hdrStrip() As String
    Dim hdrFirst() As String
    ReDim hdrStrip(1 To lastHdrCol + 64)
    ReDim hdrFirst(1 To lastHdrCol + 64)
    Dim c As Long
    For c = 1 To lastHdrCol
        Dim hRaw As String
        hRaw = CStr(shukeiWs.Cells(headerRow, c).Value)
        hdrStrip(c) = ヘッダー正規化(hRaw)
        hdrFirst(c) = Trim(Split(hRaw & vbLf, vbLf)(0))
    Next c

    ' --- BHPlan日程表からデータ収集 ---
    Dim qtyMap As Object
    Set qtyMap = CreateObject("Scripting.Dictionary")   ' "YYYY/MM/DD|列番号" -> 数量合計
    Dim dateKeys As Object
    Set dateKeys = CreateObject("Scripting.Dictionary") ' "YYYY/MM/DD" -> 日付
    Dim unmatchedKeys As Object
    Set unmatchedKeys = CreateObject("Scripting.Dictionary")
    Dim newTypeCount As Long
    newTypeCount = 0

    Dim tLastRow As Long
    tLastRow = targetWs.Cells(targetWs.Rows.Count, 1).End(xlUp).Row

    Dim i As Long
    For i = g_DataStartRow To tLastRow
        Dim model As String
        model = Trim(CStr(targetWs.Cells(i, g_ColModel).Value))
        ' V8/V9のみ（メンテは除外）
        If model <> "V8" And model <> "V9" Then GoTo NextRow

        Dim shukkaDate As Variant
        shukkaDate = targetWs.Cells(i, g_ColShukkaDate).Value
        If IsEmpty(shukkaDate) Or Not IsDate(shukkaDate) Then GoTo NextRow
        ' 当月以降のみ（過去月の集計値は保持 = 月次ローリング）
        If CDate(shukkaDate) < g_BaseDate Then GoTo NextRow

        Dim suryo As Variant
        suryo = targetWs.Cells(i, g_ColSuryo).Value
        If IsEmpty(suryo) Or Not IsNumeric(suryo) Then GoTo NextRow

        Dim dateKey As String
        dateKey = Format(CDate(shukkaDate), "YYYY/MM/DD")

        ' 3種類のキーで書込先列を解決する
        ' keys(0)=BH型式TYPE(S列): 対応列が無ければ列を新設
        ' keys(1)=LAZ型式(Y列) / keys(2)=LAY型式(Z列): 無ければ警告のみ
        Dim keys(0 To 2) As String
        keys(0) = Trim(CStr(targetWs.Cells(i, g_ColBHType).Value))
        keys(1) = Trim(CStr(targetWs.Cells(i, g_ColV8LAZType).Value))
        keys(2) = Trim(CStr(targetWs.Cells(i, g_ColV8LAYType).Value))

        Dim k As Long
        Dim qk As String
        For k = 0 To 2
            If keys(k) = "" Then GoTo NextKey
            Dim hitCount As Long
            hitCount = 0
            For c = 3 To lastHdrCol
                If 列が書込対象(hdrStrip(c), hdrFirst(c), keys(k)) Then
                    qk = dateKey & "|" & c
                    If qtyMap.Exists(qk) Then
                        qtyMap(qk) = qtyMap(qk) + CLng(suryo)
                    Else
                        qtyMap.Add qk, CLng(suryo)
                    End If
                    hitCount = hitCount + 1
                End If
            Next c
            If hitCount = 0 Then
                If k = 0 Then
                    ' 新しいBH型式 → ヘッダー末尾に列を追加
                    If lastHdrCol + 1 > UBound(hdrStrip) Then
                        ReDim Preserve hdrStrip(1 To lastHdrCol + 16)
                        ReDim Preserve hdrFirst(1 To lastHdrCol + 16)
                    End If
                    lastHdrCol = lastHdrCol + 1
                    shukeiWs.Cells(headerRow, lastHdrCol).Value = keys(0)
                    hdrStrip(lastHdrCol) = ヘッダー正規化(keys(0))
                    hdrFirst(lastHdrCol) = keys(0)
                    qk = dateKey & "|" & lastHdrCol
                    qtyMap.Add qk, CLng(suryo)
                    newTypeCount = newTypeCount + 1
                    Call ログ書込("Step13_集計表作成", "情報", "新しいBH型式列を追加: " & keys(0))
                Else
                    If Not unmatchedKeys.Exists(keys(k)) Then unmatchedKeys.Add keys(k), 1
                End If
            End If
NextKey:
        Next k

        If Not dateKeys.Exists(dateKey) Then dateKeys.Add dateKey, CDate(shukkaDate)
NextRow:
    Next i

    If qtyMap.Count = 0 Then
        Call ログ書込("Step13_集計表作成", "完了", "集計対象データなし")
        psWb.Close SaveChanges:=False
        Exit Sub
    End If

    If unmatchedKeys.Count > 0 Then
        Dim uk As Variant
        Dim ukMsg As String
        ukMsg = ""
        For Each uk In unmatchedKeys.Keys
            ukMsg = ukMsg & uk & " "
        Next uk
        Call ログ書込("Step13_集計表作成", "警告", _
            "集計シートのヘッダーに対応列が無いLAZ/LAY型式: " & ukMsg)
    End If

    ' --- 当月以降の日付行の値セルをクリア(数式セルは保持) ---
    ' 数量減少(1→0等)を反映するため。過去月の行は保持(月次ローリング)
    Dim clearedCells As Long
    clearedCells = 0
    For r = firstDataRow To lastDataRow
        If CDate(shukeiWs.Cells(r, 2).Value) >= g_BaseDate Then
            For c = 3 To lastHdrCol
                If Not shukeiWs.Cells(r, c).HasFormula Then
                    If Not IsEmpty(shukeiWs.Cells(r, c).Value) Then
                        shukeiWs.Cells(r, c).ClearContents
                        clearedCells = clearedCells + 1
                    End If
                End If
            Next c
        End If
    Next r
    Call ログ書込("Step13_集計表作成", "情報", "[DBG] クリア完了(" & clearedCells & "セル) → 行追加へ")

    ' --- 既存の日付 -> 行マップ ---
    Dim rowMap As Object
    Set rowMap = CreateObject("Scripting.Dictionary")
    For r = firstDataRow To lastDataRow
        Dim exKey As String
        exKey = Format(CDate(shukeiWs.Cells(r, 2).Value), "YYYY/MM/DD")
        If Not rowMap.Exists(exKey) Then rowMap.Add exKey, r
    Next r

    ' --- 足りない日付行を日付順の位置に挿入 ---
    Dim sortedDates() As Date
    sortedDates = 日付キー昇順(dateKeys)
    Dim addedRows As Long
    Dim tailInserted As Boolean
    addedRows = 0
    tailInserted = False

    Dim di As Long
    For di = LBound(sortedDates) To UBound(sortedDates)
        Dim d As Date
        d = sortedDates(di)
        Dim dKey As String
        dKey = Format(d, "YYYY/MM/DD")
        If rowMap.Exists(dKey) Then GoTo NextDate

        ' 挿入位置: 自分より後の日付を持つ最初の行
        Dim insertPos As Long
        insertPos = 0
        For r = firstDataRow To lastDataRow
            If CDate(shukeiWs.Cells(r, 2).Value) > d Then
                insertPos = r
                Exit For
            End If
        Next r

        Dim templateRow As Long
        If insertPos = 0 Then
            ' 既存の最終日より後 → 総計行の直上に挿入(後でSUM範囲を修復)
            insertPos = totalRow
            templateRow = lastDataRow
            tailInserted = True
        ElseIf insertPos = firstDataRow Then
            templateRow = firstDataRow
        Else
            templateRow = insertPos - 1
        End If

        ' 日付行をコピーして挿入し、数式・書式を継承する
        shukeiWs.Rows(templateRow).Copy
        shukeiWs.Rows(insertPos).Insert Shift:=xlDown
        Application.CutCopyMode = False

        ' 値セルをクリアして日付をセット(数式セルはコピーのまま)
        For c = 3 To lastHdrCol
            If Not shukeiWs.Cells(insertPos, c).HasFormula Then
                shukeiWs.Cells(insertPos, c).ClearContents
            End If
        Next c
        shukeiWs.Cells(insertPos, 2).Value = d
        If Not shukeiWs.Cells(insertPos, 1).HasFormula Then
            shukeiWs.Cells(insertPos, 1).Value = Format(d, "YY/MM")
        End If

        ' 行マップと境界を更新
        Dim rk As Variant
        For Each rk In rowMap.Keys
            If rowMap(rk) >= insertPos Then rowMap(rk) = rowMap(rk) + 1
        Next rk
        rowMap.Add dKey, insertPos
        lastDataRow = lastDataRow + 1
        totalRow = totalRow + 1
        addedRows = addedRows + 1
NextDate:
    Next di

    ' --- 総計行のSUM範囲を修復(末尾追加した場合のみ範囲外になるため) ---
    If tailInserted Then
        For c = 1 To lastHdrCol
            If shukeiWs.Cells(totalRow, c).HasFormula Then
                Dim f As String
                f = shukeiWs.Cells(totalRow, c).Formula
                If Left(UCase(f), 5) = "=SUM(" Then
                    shukeiWs.Cells(totalRow, c).Formula = _
                        "=SUM(" & 列名(c) & firstDataRow & ":" & 列名(c) & lastDataRow & ")"
                End If
            End If
        Next c
        Call ログ書込("Step13_集計表作成", "警告", _
            "既存最終日より後の日付行を追加したため総計行のSUM範囲を修復しました。" & _
            "上部の年度別ブロック(SUMIFS)の参照範囲は自動拡張されないため確認してください")
    End If

    Call ログ書込("Step13_集計表作成", "情報", "[DBG] 行追加完了(" & addedRows & "行) → 値書込へ")

    ' --- 集計値の書き込み(数式セルはスキップ=SUMIFSで自動集計) ---
    Dim writtenCount As Long
    writtenCount = 0
    Dim qkVar As Variant
    For Each qkVar In qtyMap.Keys
        Dim parts() As String
        parts = Split(CStr(qkVar), "|")
        Dim wRow As Long
        Dim wCol As Long
        wRow = rowMap(parts(0))
        wCol = CLng(parts(1))
        If Not shukeiWs.Cells(wRow, wCol).HasFormula Then
            shukeiWs.Cells(wRow, wCol).Value = qtyMap(qkVar)
            writtenCount = writtenCount + 1
        End If
    Next qkVar

    Call ログ書込("Step13_集計表作成", "情報", "[DBG] 値書込完了(" & writtenCount & "セル) → 罫線引き直し開始(" & (lastDataRow - firstDataRow + 1) & "行×" & lastHdrCol & "列)")

    ' --- 月区切り線の引き直し(全日付行) ---
    For r = firstDataRow To lastDataRow
        Dim isMonthEnd As Boolean
        If r = lastDataRow Then
            isMonthEnd = True
        Else
            Dim d1 As Date
            Dim d2 As Date
            d1 = CDate(shukeiWs.Cells(r, 2).Value)
            d2 = CDate(shukeiWs.Cells(r + 1, 2).Value)
            isMonthEnd = (Year(d1) <> Year(d2) Or Month(d1) <> Month(d2))
        End If
        For c = 1 To lastHdrCol
            With shukeiWs.Cells(r, c).Borders(xlEdgeBottom)
                If isMonthEnd Then
                    .LineStyle = xlDouble
                Else
                    .LineStyle = xlContinuous
                    .Weight = xlHairline
                End If
            End With
        Next c
    Next r

    Call ログ書込("Step13_集計表作成", "情報", "[DBG] 罫線完了 → 集計シート保存開始")
    ' 保存して閉じる
    Call 安全保存(psWb)
    Call ログ書込("Step13_集計表作成", "情報", "[DBG] 安全保存完了 → Close開始")
    On Error Resume Next
    psWb.Close SaveChanges:=False
    Dim ce As Long: ce = Err.Number
    Dim cd As String: cd = Err.Description
    Err.Clear
    On Error GoTo 0
    Call ログ書込("Step13_集計表作成", "情報", "[DBG] Close完了 Err=" & ce & " " & cd)

    Call ログ書込("Step13_集計表作成", "完了", _
        writtenCount & "セルを書込（クリア" & clearedCells & "セル、日付行追加" & addedRows & "行、" & _
        "新規型式列" & newTypeCount & "列、シート: " & shukeiWs.Name & "）")
End Sub


' ============================================================
' ヘッダーが指定キーの書込対象列かを判定する
'   ・合計/累計/出荷月 を含む列は対象外(数式列)
'   ・空白・改行を除去したヘッダーがキーと完全一致 → 対象
'     (型式別詳細列、グループ列のYB-LA001等、LAY列の"LAY(改行)001"等)
'   ・ヘッダー1行目がキーと一致 → 対象
'     (LAZグループ列: "LAZ201(改行)LC091(改行)LD091..."の1行目とY列のLAZ201)
' ============================================================
Private Function 列が書込対象(hStrip As String, hFirst As String, key As String) As Boolean
    列が書込対象 = False
    If hStrip = "" Then Exit Function
    If InStr(hStrip, "合計") > 0 Then Exit Function
    If InStr(hStrip, "累計") > 0 Then Exit Function
    If InStr(hStrip, "出荷月") > 0 Then Exit Function
    If hStrip = ヘッダー正規化(key) Then
        列が書込対象 = True
    ElseIf hFirst = Trim(key) Then
        列が書込対象 = True
    End If
End Function

' ============================================================
' ヘッダー文字列の正規化(改行・空白を除去)
' ============================================================
Private Function ヘッダー正規化(s As String) As String
    Dim t As String
    t = Replace(s, vbLf, "")
    t = Replace(t, vbCr, "")
    t = Replace(t, " ", "")
    t = Replace(t, "　", "")
    ヘッダー正規化 = Trim(t)
End Function

' ============================================================
' 列番号 -> 列名(A, B, ..., AA, ...)
' ============================================================
Private Function 列名(col As Long) As String
    列名 = Split(Cells(1, col).Address(True, False), "$")(0)
End Function

' ============================================================
' 日付キーDictionaryを昇順のDate配列にして返す
' ============================================================
Private Function 日付キー昇順(dateKeys As Object) As Date()
    Dim arr() As Date
    ReDim arr(0 To dateKeys.Count - 1)
    Dim k As Variant
    Dim n As Long
    n = 0
    For Each k In dateKeys.Keys
        arr(n) = dateKeys(k)
        n = n + 1
    Next k
    ' 挿入ソート
    Dim a As Long
    Dim b As Long
    For a = 1 To UBound(arr)
        Dim tmp As Date
        tmp = arr(a)
        b = a - 1
        Do While b >= 0
            If arr(b) <= tmp Then Exit Do
            arr(b + 1) = arr(b)
            b = b - 1
        Loop
        arr(b + 1) = tmp
    Next a
    日付キー昇順 = arr
End Function

' ============================================================
' シート名末尾が"集計"のシートを検索(0110集計, 0412集計など対応)
' g_SheetV8Shukei で見つからない場合のフォールバック
' ============================================================
Private Function 集計シート動的検索(wb As Workbook) As Worksheet
    Dim ws As Worksheet
    For Each ws In wb.Sheets
        Dim n As String
        n = Trim(ws.Name)
        If Right(n, 2) = "集計" And Len(n) >= 3 Then
            ' "XXXX集計" のような4桁数値+集計を優先
            If Len(n) = 6 Then
                If IsNumeric(Left(n, 4)) Then
                    Set 集計シート動的検索 = ws
                    Exit Function
                End If
            End If
        End If
    Next ws
    ' 末尾"集計"なら何でも(2回目走査)
    For Each ws In wb.Sheets
        If Right(Trim(ws.Name), 2) = "集計" Then
            Set 集計シート動的検索 = ws
            Exit Function
        End If
    Next ws
    Set 集計シート動的検索 = Nothing
End Function

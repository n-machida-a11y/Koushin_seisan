Attribute VB_Name = "ModStep14"
Option Explicit

' ============================================================
' Step14: 出荷・完了計画に日付行を追加
' (2026-06 正解Excel「4月の結果」に合わせて改修)
'
' V8/V9のProduction Scheduleの「出荷・完了計画」シートについて、
' 合計行(E列にSUM関数がある行)の直上に、まだ無い日付行を追加する。
'
'   ・終了日 = 「当月+3ヶ月末」と「日程表の当該MODEL最大出荷日の月末」の遅い方
'     (計画生産の全展開(2026-06)により日程表は1年先まで持つため)
'   ・追加行は直上の日付行をコピーして挿入し、数式列
'     (H:光真ストア数, I:遅れ台数, J:当日完了当日出荷, M/N/O等の累計列)
'     と書式を引き継ぐ。値セル(E/F/G、備考等)はクリアする
'   ・D列(光真稼働)は星取表計算マスターの「光真稼働日早見表」を参照
'     (祝日・会社休日対応)。早見表に無い日付は土日のみで判定
'   ・挿入は合計行の直上(SUM範囲の末尾の外側)になるため、
'     追加後に合計行のSUM範囲(E/F/G列)を修復する
' ============================================================
Public Sub Step14_出荷完了計画日付追加(targetWs As Worksheet)
    Dim defaultEnd As Date
    defaultEnd = DateSerial(Year(g_BaseDate), Month(g_BaseDate) + 4, 0)  ' 当月+3ヶ月末

    ' 日程表のMODEL別最大出荷日
    Dim maxV8 As Date
    Dim maxV9 As Date
    Call モデル別最大出荷日(targetWs, maxV8, maxV9)

    Dim endV8 As Date
    Dim endV9 As Date
    endV8 = defaultEnd
    If maxV8 > 0 Then
        If 月末(maxV8) > endV8 Then endV8 = 月末(maxV8)
    End If
    endV9 = defaultEnd
    If maxV9 > 0 Then
        If 月末(maxV9) > endV9 Then endV9 = 月末(maxV9)
    End If

    ' 光真稼働日早見表(取得できなければ土日判定で代用)
    Dim kadoubi As Object
    Dim kadoubiMax As Date
    Set kadoubi = 稼働日読み込み(kadoubiMax)

    Dim addedV8 As Long
    Dim addedV9 As Long
    addedV8 = 0
    addedV9 = 0

    If g_V8ProdSchedulePath <> "" Then
        addedV8 = 日付追加処理(g_V8ProdSchedulePath, g_SheetV8ShukkaKeikaku, endV8, kadoubi, kadoubiMax)
    End If

    If g_V9ProdSchedulePath <> "" Then
        addedV9 = 日付追加処理(g_V9ProdSchedulePath, g_SheetV9ShukkaKeikaku, endV9, kadoubi, kadoubiMax)
    End If

    Call ログ書込("Step14_出荷完了計画日付追加", "完了", _
        "V8: " & addedV8 & "日追加(～" & Format(endV8, "yyyy/mm/dd") & ")、" & _
        "V9: " & addedV9 & "日追加(～" & Format(endV9, "yyyy/mm/dd") & ")")
End Sub

' ============================================================
' 指定ファイルの出荷・完了計画シートに日付行を追加
' 合計行(E列にSUM関数がある行)の上に挿入する
' ============================================================
Private Function 日付追加処理(filePath As String, sheetName As String, endDate As Date, _
                              kadoubi As Object, kadoubiMax As Date) As Long
    日付追加処理 = 0

    Dim wb As Workbook
    On Error Resume Next
    Set wb = Workbooks.Open(filePath)
    On Error GoTo 0
    If wb Is Nothing Then
        Call ログ書込("Step14", "警告", "ファイルを開けません: " & filePath)
        Exit Function
    End If

    Dim ws As Worksheet
    Set ws = シート検索(wb, sheetName)
    If ws Is Nothing Then
        Call ログ書込("Step14", "警告", "シートが見つかりません: " & sheetName)
        wb.Close SaveChanges:=False
        Exit Function
    End If

    ' 合計行を探す（B列が空でE列とF列とG列に値がある行）
    Dim sumRow As Long
    sumRow = 0
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row

    Dim r As Long
    For r = 6 To lastRow + 20  ' 合計行はB列最終行の直後にあるため余裕を持つ
        Dim bEmpty As Boolean
        bEmpty = IsEmpty(ws.Cells(r, 2).Value) Or Trim(CStr(ws.Cells(r, 2).Value)) = ""
        If bEmpty Then
            If Not IsEmpty(ws.Cells(r, 5).Value) And _
               Not IsEmpty(ws.Cells(r, 6).Value) And _
               Not IsEmpty(ws.Cells(r, 7).Value) Then
                sumRow = r
                Exit For
            End If
        End If
    Next r

    If sumRow = 0 Then
        Call ログ書込("Step14", "警告", "合計行(SUM関数)が見つかりません: " & sheetName)
        wb.Close SaveChanges:=False
        Exit Function
    End If

    ' 既存の日付を収集（B列、合計行の手前まで）と最終日付行
    Dim existingDates As Object
    Set existingDates = CreateObject("Scripting.Dictionary")
    Dim lastDateRow As Long
    lastDateRow = 0
    For r = 6 To sumRow - 1
        Dim cellVal As Variant
        cellVal = ws.Cells(r, 2).Value
        If IsDate(cellVal) Then
            Dim dateKey As String
            dateKey = Format(CDate(cellVal), "YYYY/MM/DD")
            If Not existingDates.Exists(dateKey) Then
                existingDates.Add dateKey, r
            End If
            lastDateRow = r
        End If
    Next r
    If lastDateRow = 0 Then
        Call ログ書込("Step14", "警告", "日付行が見つかりません: " & sheetName)
        wb.Close SaveChanges:=False
        Exit Function
    End If

    Dim lastUsedCol As Long
    lastUsedCol = ws.UsedRange.Columns.Count

    ' 当月1日からendDateまで、ない日付を合計行の上に挿入
    Dim addedCount As Long
    Dim fallbackCount As Long
    addedCount = 0
    fallbackCount = 0

    Dim currentDate As Date
    currentDate = g_BaseDate

    Do While currentDate <= endDate
        dateKey = Format(currentDate, "YYYY/MM/DD")

        If Not existingDates.Exists(dateKey) Then
            ' 直上の日付行をコピーして合計行の上に挿入(数式・書式を継承)
            ws.Rows(lastDateRow).Copy
            ws.Rows(sumRow).Insert Shift:=xlDown
            Application.CutCopyMode = False

            ' 値セル(E/F/G、備考等)をクリア。数式列(H/I/J/M/N/O等)は継承
            Dim c As Long
            For c = 5 To lastUsedCol
                If Not ws.Cells(sumRow, c).HasFormula Then
                    ws.Cells(sumRow, c).ClearContents
                End If
            Next c

            ' A列: 年(=TEXT(B,"YY")等の数式ならそのまま)
            If Not ws.Cells(sumRow, 1).HasFormula Then
                ws.Cells(sumRow, 1).Value = Format(currentDate, "yy")
            End If
            ' B列: 日付
            ws.Cells(sumRow, 2).Value = currentDate
            ' C列: 曜
            ws.Cells(sumRow, 3).Value = Mid("月火水木金土日", Weekday(currentDate, vbMonday), 1)
            ' D列: 稼働日フラグ
            ws.Cells(sumRow, 4).Value = 稼働フラグ(currentDate, kadoubi, kadoubiMax, fallbackCount)

            ' 合計行が1行下にずれるので更新
            lastDateRow = sumRow
            sumRow = sumRow + 1
            addedCount = addedCount + 1
        End If

        currentDate = currentDate + 1
    Loop

    ' 合計行のSUM範囲を修復
    ' (挿入位置はSUM範囲の末尾の外側なので自動拡張されない)
    If addedCount > 0 Then
        For c = 5 To 7
            If ws.Cells(sumRow, c).HasFormula Then
                Dim f As String
                f = ws.Cells(sumRow, c).Formula
                If Left(UCase(f), 5) = "=SUM(" And InStr(f, ":") > 0 Then
                    Dim startRef As String
                    startRef = Mid(f, 6, InStr(f, ":") - 6)   ' 例: "E6"
                    Dim colLetter As String
                    colLetter = Left(startRef, 1)
                    ws.Cells(sumRow, c).Formula = _
                        "=SUM(" & startRef & ":" & colLetter & (sumRow - 1) & ")"
                End If
            End If
        Next c
    End If

    If fallbackCount > 0 Then
        Call ログ書込("Step14", "警告", sheetName & ": 稼働日早見表に無い日付" & _
            fallbackCount & "日は土日のみで稼働判定しました(祝日・会社休日は手修正してください)")
    End If

    wb.Save
    wb.Close SaveChanges:=False

    日付追加処理 = addedCount
End Function

' ============================================================
' 日程表からMODEL別(V8/V9)の最大出荷日を求める
' ============================================================
Private Sub モデル別最大出荷日(targetWs As Worksheet, ByRef maxV8 As Date, ByRef maxV9 As Date)
    maxV8 = 0
    maxV9 = 0
    Dim lastRow As Long
    lastRow = targetWs.Cells(targetWs.Rows.Count, 1).End(xlUp).Row
    Dim i As Long
    For i = g_DataStartRow To lastRow
        Dim model As String
        model = Trim(CStr(targetWs.Cells(i, g_ColModel).Value))
        If model <> "V8" And model <> "V9" Then GoTo NextRow
        Dim d As Variant
        d = targetWs.Cells(i, g_ColShukkaDate).Value
        If IsEmpty(d) Or Not IsDate(d) Then GoTo NextRow
        If model = "V8" Then
            If CDate(d) > maxV8 Then maxV8 = CDate(d)
        Else
            If CDate(d) > maxV9 Then maxV9 = CDate(d)
        End If
NextRow:
    Next i
End Sub

' ============================================================
' 月末日を返す
' ============================================================
Private Function 月末(d As Date) As Date
    月末 = DateSerial(Year(d), Month(d) + 1, 0)
End Function

' ============================================================
' 星取表計算マスターの「光真稼働日早見表」から稼働日の一覧を読み込む
' B列(当日)に載っている日付 = 稼働日
' 戻り値: Dictionary("YYYY/MM/DD" -> 1)、maxDateに早見表の最終日
' ============================================================
Private Function 稼働日読み込み(ByRef maxDate As Date) As Object
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    maxDate = 0
    Set 稼働日読み込み = dict

    If g_HoshitoriMasterPath = "" Then Exit Function

    Dim ok As Boolean
    ok = False
    On Error Resume Next
    ok = (Dir(g_HoshitoriMasterPath) <> "")
    On Error GoTo 0
    If Not ok Then
        Call ログ書込("Step14", "警告", "星取表計算マスターが見つかりません: " & g_HoshitoriMasterPath)
        Exit Function
    End If

    Dim wb As Workbook
    Set wb = Workbooks.Open(g_HoshitoriMasterPath, ReadOnly:=True)

    Dim ws As Worksheet
    Set ws = シート検索(wb, "光真稼働日早見表")
    If ws Is Nothing Then
        Call ログ書込("Step14", "警告", "光真稼働日早見表シートが見つかりません(土日判定で代用)")
        wb.Close SaveChanges:=False
        Exit Function
    End If

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row
    Dim r As Long
    For r = 2 To lastRow
        Dim v As Variant
        v = ws.Cells(r, 2).Value
        If IsDate(v) Then
            Dim k As String
            k = Format(CDate(v), "YYYY/MM/DD")
            If Not dict.Exists(k) Then dict.Add k, 1
            If CDate(v) > maxDate Then maxDate = CDate(v)
        End If
    Next r

    wb.Close SaveChanges:=False
    Call ログ書込("Step14", "情報", _
        "光真稼働日早見表を読込: " & dict.Count & "稼働日(～" & Format(maxDate, "yyyy/mm/dd") & ")")
End Function

' ============================================================
' 稼働日フラグ(○/×)を返す
' 早見表の範囲内 → 早見表に載っていれば○、無ければ×
' 早見表の範囲外/未読込 → 土日=×、平日=○(フォールバック)
' ============================================================
Private Function 稼働フラグ(d As Date, kadoubi As Object, kadoubiMax As Date, _
                            ByRef fallbackCount As Long) As String
    If Not kadoubi Is Nothing Then
        If kadoubi.Count > 0 And d <= kadoubiMax Then
            If kadoubi.Exists(Format(d, "YYYY/MM/DD")) Then
                稼働フラグ = "○"
            Else
                稼働フラグ = "×"
            End If
            Exit Function
        End If
    End If
    fallbackCount = fallbackCount + 1
    If Weekday(d, vbMonday) >= 6 Then
        稼働フラグ = "×"
    Else
        稼働フラグ = "○"
    End If
End Function

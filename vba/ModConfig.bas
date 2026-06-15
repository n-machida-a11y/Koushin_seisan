Attribute VB_Name = "ModConfig"
Option Explicit

' ===== グローバル設定変数 =====
Public g_BHPlanFolder           As String  ' BHプラン保存フォルダ
Public g_V8SavedPath            As String  ' BH計画保存版V8パス
Public g_V9SavedPath            As String  ' BH計画保存版V9パス
Public g_V8SavedKPNoCol         As Long    ' 保存版V8のKP-No列番号
Public g_V9SavedKPNoCol         As Long    ' 保存版V9のKP-No列番号
Public g_TargetSheetName        As String  ' 加工対象シート名

' 列番号
Public g_ColSeisanNo            As Long    ' B列: 生産計画No
Public g_ColKyakusakiName       As Long    ' C列: 客先名
Public g_ColKishuName           As Long    ' F列: 機種名
Public g_ColKatashiki           As Long    ' G列: 型式
Public g_ColZokusei             As Long    ' I列: 属性
Public g_ColTsuikashiyo         As Long    ' K列: 追加仕様
Public g_ColSuryo               As Long    ' L列: 数量
Public g_ColJunjoHakkoDate      As Long    ' M列: 順序指示発行日
Public g_ColShukkaDate          As Long    ' N列: 光真ss出荷日
Public g_ColKPNo                As Long    ' R列: KP-No
Public g_ColBHType              As Long    ' S列: BH型式TYPE
Public g_ColModel               As Long    ' U列: MODEL
Public g_ColKikiHinban          As Long    ' H列: 機械品番

' Phase 1-B 追加列番号
Public g_ColShipmentMonth       As Long    ' V列: shipment month
Public g_ColTNo                 As Long    ' W列: T-No
Public g_ColV8LAZType           As Long    ' Y列: V8-LAZ型式
Public g_ColV8LAYType           As Long    ' Z列: V8-LAY型式
Public g_ColZ002                As Long    ' AA列: Z002号機

' Phase 1-B ファイルパス
Public g_V8ProdSchedulePath     As String  ' V8 Production Scheduleパス
Public g_V9ProdSchedulePath     As String  ' V9 Production Scheduleパス
Public g_HoshitoriMasterPath    As String  ' 星取表計算マスターパス
Public g_BHPlan0110Path         As String  ' 0110BHPlanパス

' シート名設定
Public g_SheetV8Shukei          As String  ' V8: 0110集計シート名
Public g_SheetV8ShukkaKeikaku   As String  ' V8: 出荷・完了計画シート名
Public g_SheetV9ShukkaKeikaku   As String  ' V9: 出荷・完了計画シート名
Public g_SheetV8Hoshitori       As String  ' V8: 星取表シート名
Public g_SheetV9Hoshitori       As String  ' V9: 星取表シート名
Public g_SheetBHGraph           As String  ' BH出荷・完了グラフシート名
Public g_SheetV8KMPShipment     As String  ' V8: KMP SHIPMENT PLANシート名
Public g_SheetV9KMPShipment     As String  ' V9: KMP SHIPMENT PLANシート名
Public g_SheetKMPSchedule       As String  ' KMP出荷スケジュールシート名
Public g_SheetKMPMPS            As String  ' KMP MPSシート名
Public g_SheetV8Master          As String  ' V8星取日程マスターシート名
Public g_SheetV9Master          As String  ' V9星取表日程マスターシート名

Public g_InquiryEmail           As String  ' 問い合わせ先メール
Public g_DataStartRow           As Long    ' データ開始行番号（ヘッダー行の次の行）

' 基準日
Public g_BaseDate               As Date    ' 実行時の基準日（当月1日）

' ============================================================
' 設定シートから全設定値を読み込む
' ============================================================
Public Sub 設定読み込み()
    Dim ws As Worksheet
    Dim i As Long
    Dim key As String
    Dim val As String

    On Error GoTo ErrHandler

    Set ws = ThisWorkbook.Sheets("設定")
    g_BaseDate = DateSerial(Year(Date), Month(Date), 1)
    g_DataStartRow = 5  ' デフォルト値（設定シートで上書きされる）

    ' A列=キー, B列=値 の形式で2行目から読み込む
    For i = 2 To ws.UsedRange.Rows.Count + 1
        key = Trim(CStr(ws.Cells(i, 1).Value))
        val = Trim(CStr(ws.Cells(i, 2).Value))
        ' パス値から前後の引用符を除去（誤って"付きで入力された場合の対策）
        If Left(val, 1) = Chr(34) Then val = Mid(val, 2)
        If Right(val, 1) = Chr(34) Then val = Left(val, Len(val) - 1)
        If key = "" Then Exit For

        Select Case key
            Case "BHプラン保存フォルダ":                g_BHPlanFolder = val
            Case "BH計画保存版_V8パス":                 g_V8SavedPath = val
            Case "BH計画保存版_V9パス":                 g_V9SavedPath = val
            Case "BH計画保存版_V8_KPNo列番号":          g_V8SavedKPNoCol = CLng(val)
            Case "BH計画保存版_V9_KPNo列番号":          g_V9SavedKPNoCol = CLng(val)
            Case "加工対象シート名":                     g_TargetSheetName = val
            Case "列番号_生産計画No(B列)":               g_ColSeisanNo = CLng(val)
            Case "列番号_客先名(C列)":                   g_ColKyakusakiName = CLng(val)
            Case "列番号_機種名(F列)":                   g_ColKishuName = CLng(val)
            Case "列番号_型式(G列)":                     g_ColKatashiki = CLng(val)
            Case "列番号_追加仕様(K列)":                 g_ColTsuikashiyo = CLng(val)
            Case "列番号_数量(L列)":                     g_ColSuryo = CLng(val)
            Case "列番号_順序指示発行日(M列)":           g_ColJunjoHakkoDate = CLng(val)
            Case "列番号_光真ss出荷日(N列)":             g_ColShukkaDate = CLng(val)
            Case "列番号_KP-No(R列)":                    g_ColKPNo = CLng(val)
            Case "列番号_BH型式TYPE(S列)":               g_ColBHType = CLng(val)
            Case "列番号_MODEL(U列)":                    g_ColModel = CLng(val)
            Case "列番号_属性(I列)":                     g_ColZokusei = CLng(val)
            Case "列番号_機械品番(H列)":                 g_ColKikiHinban = CLng(val)
            Case "問い合わせ先メール":                    g_InquiryEmail = val
            Case "データ開始行番号":                      g_DataStartRow = CLng(val)
            ' Phase 1-B 追加設定
            Case "列番号_shipment month(V列)":           g_ColShipmentMonth = CLng(val)
            Case "列番号_T-No(W列)":                     g_ColTNo = CLng(val)
            Case "列番号_V8-LAZ型式(Y列)":               g_ColV8LAZType = CLng(val)
            Case "列番号_V8-LAY型式(Z列)":               g_ColV8LAYType = CLng(val)
            Case "列番号_Z002号機(AA列)":                g_ColZ002 = CLng(val)
            Case "V8_ProductionScheduleパス":            g_V8ProdSchedulePath = val
            Case "V9_ProductionScheduleパス":            g_V9ProdSchedulePath = val
            Case "星取表計算マスターパス":                g_HoshitoriMasterPath = val
            Case "0110BHPlanパス":                       g_BHPlan0110Path = val
            ' シート名設定
            Case "シート名_V8_0110集計":                 g_SheetV8Shukei = val
            Case "シート名_V8_出荷完了計画":             g_SheetV8ShukkaKeikaku = val
            Case "シート名_V9_出荷完了計画":             g_SheetV9ShukkaKeikaku = val
            Case "シート名_V8_星取表":                   g_SheetV8Hoshitori = val
            Case "シート名_V9_星取表":                   g_SheetV9Hoshitori = val
            Case "シート名_BH出荷完了グラフ":            g_SheetBHGraph = val
            Case "シート名_V8_KMP_SHIPMENT":             g_SheetV8KMPShipment = val
            Case "シート名_V9_KMP_SHIPMENT":             g_SheetV9KMPShipment = val
            Case "シート名_KMP出荷スケジュール":         g_SheetKMPSchedule = val
            Case "シート名_KMP_MPS":                     g_SheetKMPMPS = val
            Case "シート名_V8星取日程マスター":           g_SheetV8Master = val
            Case "シート名_V9星取表日程マスター":         g_SheetV9Master = val
        End Select
    Next i

    Exit Sub
ErrHandler:
    MsgBox "設定シートの読み込みに失敗しました。" & vbCrLf & _
           "設定シートの内容を確認してください。" & vbCrLf & _
           "エラー: " & Err.Description, vbCritical, "設定読み込みエラー"
    End
End Sub


' ============================================================
' ワークブックを安全に保存する
' 計算が手動(xlCalculationManual)のまま、複合グラフ
' (barChart+lineChart 等)を含むブックを保存すると Excel が
' 「指定したディメンションは、このグラフの種類では無効です」
' エラーを出すため、保存時だけ自動計算に戻して保存する。
' 万一それでもエラーが出た場合は停止せず警告ログを残して続行する。
' (2026-06-15 実機テストで Production schedule 保存時に発生)
' 戻り値: 保存成功=True / 警告あり=False
' ============================================================
' ============================================================
' ブック内の全グラフから、壊れた参照(#REF!)を含む系列を除去する
' 複合グラフ(barChart+lineChart)に #REF! を含む系列があると、
' マクロ実行中(画面更新オフ・計算手動)の保存で
' 「指定したディメンションは、このグラフの種類では無効です」
' エラーになる。手作業の保存では出ないがマクロ保存のみ発生する。
' (2026-06-15 実機: グラフ削除でエラー消失と切り分け済み)
' ============================================================
Public Sub グラフ壊れ系列除去(wb As Workbook)
    On Error Resume Next
    Dim removed As Long
    removed = 0
    Dim ws As Worksheet
    For Each ws In wb.Worksheets
        Dim co As ChartObject
        For Each co In ws.ChartObjects
            Dim si As Long
            For si = co.Chart.SeriesCollection.Count To 1 Step -1
                Dim f As String
                f = ""
                f = co.Chart.SeriesCollection(si).Formula
                If InStr(f, "#REF!") > 0 Then
                    co.Chart.SeriesCollection(si).Delete
                    removed = removed + 1
                End If
            Next si
        Next co
    Next ws
    On Error GoTo 0
    If removed > 0 Then
        Call ログ書込("グラフ修復", "情報", _
            wb.Name & " のグラフから壊れた系列(#REF!)を " & removed & " 本除去しました")
    End If
End Sub

Public Function 安全保存(wb As Workbook) As Boolean
    ' === 切り分け用: 各段階でErr番号をログに出す ===
    Dim e1 As Long, e2 As Long, e3 As Long, e4 As Long
    Dim d3 As String

    On Error Resume Next

    ' (1) グラフの壊れた系列(#REF!)を除去
    Call グラフ壊れ系列除去(wb)
    e1 = Err.Number: Err.Clear
    Call ログ書込("安全保存", "情報", "[DBG] (1)グラフ除去後 Err=" & e1)

    ' (2) 計算を自動に戻す(手作業と同じ状態)
    Dim prevCalc As Long
    prevCalc = Application.Calculation
    Application.Calculation = xlCalculationAutomatic
    e2 = Err.Number: Err.Clear
    Call ログ書込("安全保存", "情報", "[DBG] (2)計算Auto化後 Err=" & e2)

    ' (3) 保存本体
    wb.Save
    e3 = Err.Number: d3 = Err.Description: Err.Clear
    Call ログ書込("安全保存", "情報", "[DBG] (3)wb.Save後 Err=" & e3 & " " & d3)

    ' (4) 計算モードを元に戻す
    Application.Calculation = prevCalc
    e4 = Err.Number: Err.Clear
    Call ログ書込("安全保存", "情報", "[DBG] (4)計算戻し後 Err=" & e4)

    On Error GoTo 0

    安全保存 = (e3 = 0)
End Function

' ============================================================
' シート名でシートを取得（完全一致→部分一致のフォールバック）
' 末尾スペース等の微妙な違いに対応
' ============================================================
Public Function シート検索(wb As Workbook, sheetName As String) As Worksheet
    ' 1. 完全一致
    On Error Resume Next
    Set シート検索 = wb.Sheets(sheetName)
    On Error GoTo 0
    If Not シート検索 Is Nothing Then Exit Function
    
    ' 2. Trim一致
    Dim ws As Worksheet
    For Each ws In wb.Sheets
        If Trim(ws.Name) = Trim(sheetName) Then
            Set シート検索 = ws
            Exit Function
        End If
    Next ws
    
    ' 3. 部分一致（sheetNameを含むシート）
    For Each ws In wb.Sheets
        If InStr(ws.Name, sheetName) > 0 Then
            Set シート検索 = ws
            Exit Function
        End If
    Next ws
    
    Set シート検索 = Nothing
End Function

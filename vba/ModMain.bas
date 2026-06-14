Attribute VB_Name = "ModMain"
Option Explicit

' ============================================================
' メイン実行: Phase 1-A（Step5～12）
' 「設定」シートのボタンから呼び出す
' ============================================================
Public Sub メイン実行()
    Dim ans As VbMsgBoxResult
    ans = MsgBox("生産計画自動化（Phase 1-A）を開始します。" & vbCrLf & vbCrLf & _
                 "【事前確認】" & vbCrLf & _
                 "・BHプランの出力ファイル（xlsx）を BHプラン保存フォルダに置いてください" & vbCrLf & _
                 "・設定シートのフォルダパスが正しいことを確認してください" & vbCrLf & vbCrLf & _
                 "続行しますか？", vbYesNo + vbQuestion, "生産計画自動化")
    If ans = vbNo Then Exit Sub

    Application.AskToUpdateLinks = False
    Call 設定読み込み

    Dim targetWb As Workbook
    Set targetWb = 対象ファイルを開く()
    If targetWb Is Nothing Then Exit Sub

    Dim targetWs As Worksheet
    On Error Resume Next
    Set targetWs = targetWb.Sheets(g_TargetSheetName)
    On Error GoTo 0
    If targetWs Is Nothing Then
        MsgBox "シート「" & g_TargetSheetName & "」が見つかりません。" & vbCrLf & _
               "設定シートの「加工対象シート名」を確認してください。", _
               vbCritical, "シートが見つかりません"
        targetWb.Close SaveChanges:=False
        Exit Sub
    End If

    Call ログ書込("メイン実行", "情報", "処理開始: " & targetWb.Name)

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    On Error GoTo ErrHandler

    ' ===== Step5～12 =====
    ' 2026-06ヒアリング指摘【A】: 展開(Step08)を出荷済み削除(Step06)より先に実行。
    ' V9はKP-Noが無く、星取表には枝番付き生産計画No(-01等)で記録されているため、
    ' 先に展開して枝番を付与しないと出荷済み行を紐付けて削除できない。
    Call Step05_計画生産対象削除(targetWs)
    Call Step08_計画生産行展開(targetWs)
    Call Step06_出荷済みデータ削除(targetWs)
    Call Step07_型式補完(targetWs)
    Call Step09_数量チェック(targetWs)
    Call Step10_並び替え(targetWs)
    Call Step11_メンテ照合(targetWs)
    Call Step12_連番付与(targetWs)

    On Error GoTo 0

    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True

    targetWb.Save

    Call ログ書込("メイン実行", "成功", "Phase 1-A 処理完了: " & targetWb.Name)
    MsgBox "Phase 1-A 完了（データ加工）" & vbCrLf & vbCrLf & _
           "「ログ」シートで結果を確認してください。" & vbCrLf & _
           "問題なければ Phase 1-B1 を実行してください。", vbInformation, "完了"
    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True

    Call ログ書込("メイン実行", "エラー", "予期しないエラー: " & Err.Description)
    MsgBox "予期しないエラーが発生しました。" & vbCrLf & _
           "ファイルは保存されていません。" & vbCrLf & vbCrLf & _
           "エラー: " & Err.Description, vbCritical, "エラー"
    targetWb.Close SaveChanges:=False
End Sub

' ============================================================
' Phase 1-B1 実行: 集計・転記（Step13～20）
' ============================================================
Public Sub Phase1B1実行()
    Application.AskToUpdateLinks = False
    Call 設定読み込み

    Dim targetWb As Workbook
    Set targetWb = 対象ファイルを開く()
    If targetWb Is Nothing Then Exit Sub

    Dim targetWs As Worksheet
    On Error Resume Next
    Set targetWs = targetWb.Sheets(g_TargetSheetName)
    On Error GoTo 0
    If targetWs Is Nothing Then
        MsgBox "シート「" & g_TargetSheetName & "」が見つかりません。", vbCritical
        targetWb.Close SaveChanges:=False
        Exit Sub
    End If

    Call ログ書込("Phase1B1実行", "情報", "集計・転記開始: " & targetWb.Name)

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    On Error GoTo ErrB1

    Call ログ書込("Phase1B1実行", "情報", "[DBG] Step13呼出 前")
    Call Step13_集計表作成(targetWs)
    Call ログ書込("Phase1B1実行", "情報", "[DBG] Step13完了 → Step14呼出 前")
    Call Step14_出荷完了計画日付追加(targetWs)
    Call ログ書込("Phase1B1実行", "情報", "[DBG] Step14完了 → Step15呼出 前")
    Call Step15_出荷台数入力(targetWs)
    Call ログ書込("Phase1B1実行", "情報", "[DBG] Step15完了")

    On Error GoTo 0

    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True

    Call ログ書込("Phase1B1実行", "成功", "Phase 1-B1 処理完了")
    MsgBox "Phase 1-B1 完了（集計・日付・台数入力）" & vbCrLf & vbCrLf & _
           "【次の作業】" & vbCrLf & _
           "出荷・完了計画のF列（光真完了台数）に" & vbCrLf & _
           "平準化した台数を手作業で入力してください。" & vbCrLf & vbCrLf & _
           "平準化が完了したら Phase 1-B2 を実行してください。", vbInformation, "完了"
    Exit Sub

ErrB1:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Call ログ書込("Phase1B1実行", "エラー", "予期しないエラー: " & Err.Description)
    MsgBox "予期しないエラーが発生しました。" & vbCrLf & _
           "エラー: " & Err.Description, vbCritical, "エラー"
End Sub

' ============================================================
' Phase 1-B2 実行: マスター更新・KMP関連（Step17～26）
' ============================================================
Public Sub Phase1B2実行()
    Application.AskToUpdateLinks = False
    Call 設定読み込み

    Dim targetWb As Workbook
    Set targetWb = 対象ファイルを開く()
    If targetWb Is Nothing Then Exit Sub

    Dim targetWs As Worksheet
    On Error Resume Next
    Set targetWs = targetWb.Sheets(g_TargetSheetName)
    On Error GoTo 0
    If targetWs Is Nothing Then
        MsgBox "シート「" & g_TargetSheetName & "」が見つかりません。", vbCritical
        targetWb.Close SaveChanges:=False
        Exit Sub
    End If

    Call ログ書込("Phase1B2実行", "情報", "マスター更新・KMP関連開始: " & targetWb.Name)

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    On Error GoTo ErrB2

    Call Step17_V8マスター更新(targetWs)
    Call Step18_V9マスター更新(targetWs)
    Call Step19_星取表差替(targetWs)
    Call Step20_グラフ更新(targetWs)
    Call Step21_KMPユニット必要台数(targetWs)
    Call Step22_KMPシップメント更新(targetWs)
    Call Step24_KMP出荷スケジュール更新(targetWs)
    Call Step25_KMP_MPS更新(targetWs)
    Call Step26_Forecast完成(targetWs)

    On Error GoTo 0

    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True

    Call ログ書込("Phase1B2実行", "成功", "Phase 1-B2 処理完了")
    MsgBox "Phase 1-B2 完了（マスター更新・KMP関連）" & vbCrLf & vbCrLf & _
           "「ログ」シートで結果を確認してください。" & vbCrLf & _
           "全フェーズの処理が完了しました。", vbInformation, "完了"
    Exit Sub

ErrB2:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Call ログ書込("Phase1B2実行", "エラー", "予期しないエラー: " & Err.Description)
    MsgBox "予期しないエラーが発生しました。" & vbCrLf & _
           "エラー: " & Err.Description, vbCritical, "エラー"
End Sub

' ============================================================
' BHプラン保存フォルダ内の xlsx ファイルを開いて返す
' 複数ある場合は最後に更新されたものを選ぶ
' ============================================================
Private Function 対象ファイルを開く() As Workbook
    Dim folderPath As String
    folderPath = g_BHPlanFolder
    If Right(folderPath, 1) <> "\" Then folderPath = folderPath & "\"

    Dim fileName As String
    Dim latestFile As String
    Dim latestDate As Date
    Dim dirErrNum As Long
    dirErrNum = 0
    On Error Resume Next
    fileName = Dir(folderPath & "*.xlsx")
    dirErrNum = Err.Number
    On Error GoTo 0
    If dirErrNum <> 0 Then
        MsgBox "フォルダへのアクセスに失敗しました(Error " & dirErrNum & ")。" & vbCrLf & _
               "フォルダ: " & folderPath & vbCrLf & _
               "設定シートの「BHプラン保存フォルダ」を確認してください。", _
               vbCritical, "フォルダアクセスエラー"
        Set 対象ファイルを開く = Nothing
        Exit Function
    End If

    Do While fileName <> ""
        Dim fileDate As Date
        fileDate = FileDateTime(folderPath & fileName)
        If fileDate > latestDate Then
            latestDate = fileDate
            latestFile = fileName
        End If
        fileName = Dir()
    Loop

    If latestFile = "" Then
        MsgBox "BHプラン保存フォルダにxlsxファイルが見つかりません。" & vbCrLf & _
               "フォルダ: " & folderPath & vbCrLf & vbCrLf & _
               "BHプランの出力ファイルをフォルダに配置してから再実行してください。", _
               vbCritical, "ファイルなし"
        Set 対象ファイルを開く = Nothing
        Exit Function
    End If

    Set 対象ファイルを開く = Workbooks.Open(folderPath & latestFile)
End Function

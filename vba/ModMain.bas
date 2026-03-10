Attribute VB_Name = "ModMain"
Option Explicit

' ============================================================
' メイン実行: Phase 1-A（ステップ⑤〜⑩）
' 「設定」シートのボタンから呼び出す
' ============================================================
Public Sub メイン実行()
    ' 開始確認ダイアログ
    Dim ans As VbMsgBoxResult
    ans = MsgBox("生産計画自動化（Phase 1-A）を開始します。" & vbCrLf & vbCrLf & _
                 "【事前確認】" & vbCrLf & _
                 "・BHプランの出力ファイル（xlsx）を inputフォルダに置いてください" & vbCrLf & _
                 "・設定シートのフォルダパスが正しいことを確認してください" & vbCrLf & vbCrLf & _
                 "続行しますか？", vbYesNo + vbQuestion, "生産計画自動化")
    If ans = vbNo Then Exit Sub

    ' 設定読み込み
    Call 設定読み込み()

    ' 加工対象ファイルを開く
    Dim targetWb As Workbook
    Set targetWb = 対象ファイルを開く()
    If targetWb Is Nothing Then Exit Sub

    ' 対象シートを取得
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

    ' ログに開始を記録
    Call ログ書込("メイン実行", "情報", "処理開始: " & targetWb.Name)

    ' ===== ステップ⑤〜⑩を順番に実行 =====
    Call Step05_計画生産対象削除(targetWs)
    Call Step06_出荷済みデータ削除(targetWs)
    Call Step07_型式補完(targetWs)
    Call Step08_計画生産行展開(targetWs)
    Call Step09_数量チェック(targetWs)
    Call Step10_並び替え(targetWs)
    ' =========================================

    ' 上書き保存
    targetWb.Save

    Call ログ書込("メイン実行", "成功", "Phase 1-A 処理完了: " & targetWb.Name)
    MsgBox "処理が完了しました。" & vbCrLf & _
           "「ログ」シートで処理結果を確認してください。", vbInformation, "完了"
End Sub

' ============================================================
' inputフォルダ内の xlsx ファイルを開いて返す
' 複数ある場合は最後に更新されたものを選ぶ
' ============================================================
Private Function 対象ファイルを開く() As Workbook
    Dim folderPath As String
    folderPath = g_BHPlanFolder
    If Right(folderPath, 1) <> "\" Then folderPath = folderPath & "\"

    ' フォルダ内のxlsxを検索して最新ファイルを取得
    Dim fileName As String
    Dim latestFile As String
    Dim latestDate As Date
    fileName = Dir(folderPath & "*.xlsx")

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
        MsgBox "inputフォルダにxlsxファイルが見つかりません。" & vbCrLf & _
               "フォルダ: " & folderPath & vbCrLf & vbCrLf & _
               "BHプランの出力ファイルをフォルダに配置してから再実行してください。", _
               vbCritical, "ファイルなし"
        Set 対象ファイルを開く = Nothing
        Exit Function
    End If

    Set 対象ファイルを開く = Workbooks.Open(folderPath & latestFile)
End Function

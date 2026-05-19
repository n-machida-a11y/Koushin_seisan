Attribute VB_Name = "ModLoader"
Option Explicit

' ============================================================
' VBA一括インポート用ローダー
'
' 使い方:
'   1. このブックを開いた状態で Alt+F11 でVBE起動
'   2. ModLoader を選択し F5 で VBA一括インポート を実行
'   3. 完了ダイアログで「成功: 25件」を確認
'   4. Ctrl+S でブック保存
'
' 前提:
'   ファイル→オプション→トラストセンター→トラストセンターの設定→
'   マクロの設定 で「VBAプロジェクトオブジェクトモデルへのアクセスを
'   信頼する」にチェックを入れておくこと
' ============================================================

' ============================================================
' vba\\ フォルダの全 .bas ファイルを一括インポート
' 既存の同名モジュールは削除してから取り込む
' ============================================================
Public Sub VBA一括インポート()
    Dim basFolder As String
    basFolder = ThisWorkbook.Path & "\\vba\\"

    If Dir(basFolder, vbDirectory) = "" Then
        MsgBox "vba フォルダが見つかりません: " & basFolder & vbCrLf & _
               "ブックと vba フォルダを同じ親フォルダに置いてください。", _
               vbCritical, "ModLoader"
        Exit Sub
    End If

    ' インポート順 (基盤 → Step → Main)
    Dim modules As Variant
    modules = Array( _
        "ModConfig", "ModLog", "ModError", "ModDebug", _
        "ModStep05", "ModStep06", "ModStep07", "ModStep08", "ModStep09", _
        "ModStep10", "ModStep11", "ModStep12", "ModStep13", "ModStep14", _
        "ModStep15", "ModStep17", "ModStep18", "ModStep19", "ModStep20", _
        "ModStep21", "ModStep22", "ModStep24", "ModStep25", "ModStep26", _
        "ModMain" _
    )

    Dim vbProj As Object
    On Error Resume Next
    Set vbProj = ThisWorkbook.VBProject
    If Err.Number <> 0 Then
        On Error GoTo 0
        MsgBox "VBAプロジェクトにアクセスできません。" & vbCrLf & vbCrLf & _
               "[対処] ファイル→オプション→トラストセンター→" & vbCrLf & _
               "トラストセンターの設定→マクロの設定 で" & vbCrLf & _
               "「VBAプロジェクトオブジェクトモデルへのアクセスを信頼する」" & vbCrLf & _
               "にチェックを入れてください。", vbCritical, "ModLoader"
        Exit Sub
    End If
    On Error GoTo 0

    Dim importedCount As Long
    Dim skippedCount As Long
    Dim errors As String
    importedCount = 0
    skippedCount = 0
    errors = ""

    Dim modName As Variant
    For Each modName In modules
        Dim basPath As String
        basPath = basFolder & CStr(modName) & ".bas"

        If Dir(basPath) = "" Then
            skippedCount = skippedCount + 1
            errors = errors & "  ・" & CStr(modName) & ".bas が見つかりません" & vbCrLf
            GoTo NextMod
        End If

        ' 既存モジュールを削除 (なければスキップ)
        On Error Resume Next
        Dim existing As Object
        Set existing = vbProj.VBComponents(CStr(modName))
        If Err.Number = 0 Then
            vbProj.VBComponents.Remove existing
        End If
        Err.Clear
        On Error GoTo 0

        ' インポート
        On Error Resume Next
        vbProj.VBComponents.Import basPath
        If Err.Number <> 0 Then
            errors = errors & "  ・" & CStr(modName) & ": " & Err.Description & vbCrLf
            skippedCount = skippedCount + 1
        Else
            importedCount = importedCount + 1
        End If
        On Error GoTo 0
NextMod:
    Next modName

    Dim msg As String
    msg = "VBA一括インポート完了" & vbCrLf & vbCrLf & _
          "  成功: " & importedCount & " 件" & vbCrLf & _
          "  スキップ/失敗: " & skippedCount & " 件"
    If errors <> "" Then msg = msg & vbCrLf & vbCrLf & "詳細:" & vbCrLf & errors
    msg = msg & vbCrLf & vbCrLf & _
          "Ctrl+S でブックを保存してください。"

    MsgBox msg, IIf(skippedCount = 0, vbInformation, vbExclamation), "ModLoader"
End Sub

' ============================================================
' 既存の標準モジュールを全削除 (クリーンインポート用)
' ThisWorkbook / シート / ModLoader 自身は残す
' ============================================================
Public Sub VBA全モジュール削除()
    Dim ans As VbMsgBoxResult
    ans = MsgBox("ModLoader以外の全標準モジュールを削除します。" & vbCrLf & _
                 "クリーンな状態でインポートし直したい時に使います。" & vbCrLf & vbCrLf & _
                 "続行しますか？", _
                 vbYesNo + vbExclamation, "ModLoader: 全モジュール削除")
    If ans = vbNo Then Exit Sub

    Dim vbProj As Object
    On Error Resume Next
    Set vbProj = ThisWorkbook.VBProject
    If Err.Number <> 0 Then
        On Error GoTo 0
        MsgBox "VBAプロジェクトにアクセスできません。" & vbCrLf & _
               "トラストセンターの設定を確認してください。", vbCritical
        Exit Sub
    End If
    On Error GoTo 0

    Dim deletedCount As Long
    deletedCount = 0

    Dim i As Long
    For i = vbProj.VBComponents.Count To 1 Step -1
        Dim comp As Object
        Set comp = vbProj.VBComponents(i)
        ' Type 1 = 標準モジュール のみ削除
        If comp.Type = 1 And comp.Name <> "ModLoader" Then
            vbProj.VBComponents.Remove comp
            deletedCount = deletedCount + 1
        End If
    Next i

    MsgBox deletedCount & " 件のモジュールを削除しました。" & vbCrLf & _
           "続けて VBA一括インポート を実行してください。", _
           vbInformation, "ModLoader"
End Sub

# 生産計画自動化 Phase 1-A 実装計画

> **For agentic workers:** REQUIRED: Use superpowers:subagent-driven-development (if subagents available) or superpowers:executing-plans to implement this plan. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** `生産計画_マクロ.xlsm` を新規作成し、設定シート群と生産計画データのコア加工処理（ステップ⑤〜⑩）をVBAマクロとして実装する。

**Architecture:** `生産計画_マクロ.xlsm` 1ファイルに設定シート群（設定・列対応表・稼働日カレンダー・ログ）とVBAモジュール群を格納する。設定値はすべてシートから読み込み、コード内にハードコードしない。各ステップは独立したSubルーチンとして実装し、`メイン実行()`から順番に呼び出す。

**Tech Stack:** Excel VBA (xlsm), openpyxl (ファイル初期作成用Python), Windows Excel

---

## ファイル構成

```
Pドライブ/生産計画自動化/
  生産計画_マクロ.xlsm
    │
    ├─ [Sheet] 設定          ← ファイルパス・列番号など変更可能な設定値
    ├─ [Sheet] 列対応表       ← 機種ごとの列範囲・ファイルパス対応表
    ├─ [Sheet] 稼働日カレンダー ← 自社・KMP稼働日フラグ
    ├─ [Sheet] ログ           ← 実行履歴・警告の自動記録
    │
    ├─ [Module] ModConfig     ← 設定値読み込み・グローバル変数定義
    ├─ [Module] ModMain       ← メイン実行・ステップ呼び出し
    ├─ [Module] ModLog        ← ログ書き込みユーティリティ
    ├─ [Module] ModError      ← エラーハンドリング・ハイライト・ポップアップ
    ├─ [Module] ModStep05     ← ステップ⑤: 計画生産対象削除
    ├─ [Module] ModStep06     ← ステップ⑥: 出荷済みデータ削除
    ├─ [Module] ModStep07     ← ステップ⑦: 型式（S列）補完
    ├─ [Module] ModStep08     ← ステップ⑧: 計画生産行展開（1台1行化）
    ├─ [Module] ModStep09     ← ステップ⑨: 数量チェック（V8/V9 3ヶ月以内）
    └─ [Module] ModStep10     ← ステップ⑩: 並び替え

input/                        ← BHプランなど入力ファイル置き場（手動で配置）
```

---

## Chunk 1: ファイル初期作成と設定シート群

### Task 1: `生産計画_マクロ.xlsm` の雛形を作成する

**Files:**
- Create: `生産計画_マクロ.xlsm`（Python/openpyxlで生成後、以降VBAで編集）

- [ ] **Step 1: 作業ディレクトリを確認する**

```bash
ls /home/user/koushin_seisan/
```
期待: `メモつきOSS生産計画作成手順.xlsx` などが存在すること

- [ ] **Step 2: Pythonで雛形xlsmファイルを作成するスクリプトを書く**

ファイル: `scripts/create_template.py`

```python
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

wb = Workbook()

# シート1: 設定
ws_config = wb.active
ws_config.title = "設定"

headers = ["項目名", "値", "備考"]
ws_config.append(headers)
for col in range(1, 4):
    ws_config.cell(1, col).font = Font(bold=True)

rows = [
    ["BHプラン保存フォルダ", "P:\\生産計画\\input\\", "BHプランを置くフォルダ"],
    ["BH計画保存版_V8パス", "P:\\保存版\\V8_BH計画保存版.xlsx", ""],
    ["BH計画保存版_V9パス", "P:\\保存版\\V9_BH計画保存版.xlsx", ""],
    ["加工対象シート名", "日程表", "光真システムから出力されたシート名"],
    ["列番号_生産計画No(B列)", "2", ""],
    ["列番号_客先名(C列)", "3", ""],
    ["列番号_機種名(F列)", "6", ""],
    ["列番号_型式(G列)", "7", ""],
    ["列番号_追加仕様(K列)", "11", ""],
    ["列番号_数量(L列)", "12", ""],
    ["列番号_順序指示発行日(M列)", "13", ""],
    ["列番号_光真ss出荷日(N列)", "14", ""],
    ["列番号_KP-No(R列)", "18", ""],
    ["列番号_BH型式TYPE(S列)", "19", ""],
    ["列番号_MODEL(U列)", "21", ""],
    ["列番号_属性(I列)", "9", ""],
    ["列番号_機械品番(H列)", "8", ""],
    ["問い合わせ先メール", "", "オムロン担当者メールアドレス"],
]
for row in rows:
    ws_config.append(row)

ws_config.column_dimensions["A"].width = 30
ws_config.column_dimensions["B"].width = 40
ws_config.column_dimensions["C"].width = 30

# シート2: 列対応表
ws_cols = wb.create_sheet("列対応表")
ws_cols.append(["機種", "MODEL値", "保存版ファイル設定キー", "備考"])
for col in range(1, 5):
    ws_cols.cell(1, col).font = Font(bold=True)
ws_cols.append(["V8", "V8", "BH計画保存版_V8パス", ""])
ws_cols.append(["V9", "V9", "BH計画保存版_V9パス", ""])
ws_cols.append(["メンテV8", "ﾒﾝﾃV8", "BH計画保存版_V8パス", ""])
ws_cols.append(["メンテV9", "ﾒﾝﾃV9", "BH計画保存版_V9パス", ""])
for col_letter in ["A", "B", "C", "D"]:
    ws_cols.column_dimensions[col_letter].width = 25

# シート3: 稼働日カレンダー
ws_cal = wb.create_sheet("稼働日カレンダー")
ws_cal.append(["日付", "自社稼働(○/×)", "KMP稼働(○/×)", "備考"])
for col in range(1, 5):
    ws_cal.cell(1, col).font = Font(bold=True)
ws_cal.column_dimensions["A"].width = 15
ws_cal.column_dimensions["B"].width = 18
ws_cal.column_dimensions["C"].width = 18
ws_cal.column_dimensions["D"].width = 20

# シート4: ログ
ws_log = wb.create_sheet("ログ")
ws_log.append(["実行日時", "ステップ", "結果", "メッセージ"])
for col in range(1, 5):
    ws_log.cell(1, col).font = Font(bold=True)
ws_log.column_dimensions["A"].width = 20
ws_log.column_dimensions["B"].width = 20
ws_log.column_dimensions["C"].width = 12
ws_log.column_dimensions["D"].width = 60

wb.save("/home/user/koushin_seisan/生産計画_マクロ.xlsm")
print("作成完了: 生産計画_マクロ.xlsm")
```

- [ ] **Step 3: スクリプトを実行してファイルを生成する**

```bash
python3 scripts/create_template.py
```
期待: `生産計画_マクロ.xlsm` が作成される

- [ ] **Step 4: 生成されたファイルをExcelで開いて設定シートの構造を確認する**

確認項目:
- 「設定」シートに全行が存在する
- 「列対応表」「稼働日カレンダー」「ログ」シートが存在する
- ヘッダーが太字になっている

- [ ] **Step 5: コミット**

```bash
cd /home/user/koushin_seisan
git add scripts/create_template.py 生産計画_マクロ.xlsm
git commit -m "feat: 生産計画_マクロ.xlsm 雛形と設定シート群を作成"
```

---

## Chunk 2: VBAモジュール基盤（設定読み込み・ログ・エラー処理）

### Task 2: ModConfig — 設定読み込みモジュール

**Files:**
- VBAモジュール `ModConfig` を `生産計画_マクロ.xlsm` に追加

**概要:** 「設定」シートから全設定値を読み込み、グローバル変数として保持する。

- [ ] **Step 1: Excelの VBAエディタ（Alt+F11）を開き、新規モジュール `ModConfig` を挿入して以下のコードを貼り付ける**

```vba
Option Explicit

' ===== グローバル設定変数 =====
Public g_BHPlanFolder       As String  ' BHプラン保存フォルダ
Public g_V8SavedPath        As String  ' BH計画保存版V8パス
Public g_V9SavedPath        As String  ' BH計画保存版V9パス
Public g_TargetSheetName    As String  ' 加工対象シート名

' 列番号
Public g_ColSeisanNo        As Long    ' B列: 生産計画No
Public g_ColKyakusakiName   As Long    ' C列: 客先名
Public g_ColKishuName       As Long    ' F列: 機種名
Public g_ColKatashiki       As Long    ' G列: 型式
Public g_ColTsuikashiyo     As Long    ' K列: 追加仕様
Public g_ColSuryo           As Long    ' L列: 数量
Public g_ColJunjoHakkoDate  As Long    ' M列: 順序指示発行日
Public g_ColShukkaDate      As Long    ' N列: 光真ss出荷日
Public g_ColKPNo            As Long    ' R列: KP-No
Public g_ColBHType          As Long    ' S列: BH型式TYPE
Public g_ColModel           As Long    ' U列: MODEL
Public g_ColZokusei         As Long    ' I列: 属性
Public g_ColKikiHinban      As Long    ' H列: 機械品番

Public g_InquiryEmail       As String  ' 問い合わせ先メール

' 基準日
Public g_BaseDate           As Date    ' 実行時の基準日（当月1日）

' 設定シートから全設定値を読み込む
Public Sub 設定読み込み()
    Dim ws As Worksheet
    Dim i As Long
    Dim key As String
    Dim val As String

    On Error GoTo ErrHandler

    ws = ThisWorkbook.Sheets("設定")
    g_BaseDate = DateSerial(Year(Date), Month(Date), 1)

    ' A列=キー, B列=値 の形式で読み込む（2行目から）
    For i = 2 To ws.UsedRange.Rows.Count + 1
        key = Trim(ws.Cells(i, 1).Value)
        val = Trim(ws.Cells(i, 2).Value)
        If key = "" Then Exit For

        Select Case key
            Case "BHプラン保存フォルダ":        g_BHPlanFolder = val
            Case "BH計画保存版_V8パス":         g_V8SavedPath = val
            Case "BH計画保存版_V9パス":         g_V9SavedPath = val
            Case "加工対象シート名":             g_TargetSheetName = val
            Case "列番号_生産計画No(B列)":       g_ColSeisanNo = CLng(val)
            Case "列番号_客先名(C列)":           g_ColKyakusakiName = CLng(val)
            Case "列番号_機種名(F列)":           g_ColKishuName = CLng(val)
            Case "列番号_型式(G列)":             g_ColKatashiki = CLng(val)
            Case "列番号_追加仕様(K列)":         g_ColTsuikashiyo = CLng(val)
            Case "列番号_数量(L列)":             g_ColSuryo = CLng(val)
            Case "列番号_順序指示発行日(M列)":   g_ColJunjoHakkoDate = CLng(val)
            Case "列番号_光真ss出荷日(N列)":     g_ColShukkaDate = CLng(val)
            Case "列番号_KP-No(R列)":            g_ColKPNo = CLng(val)
            Case "列番号_BH型式TYPE(S列)":       g_ColBHType = CLng(val)
            Case "列番号_MODEL(U列)":            g_ColModel = CLng(val)
            Case "列番号_属性(I列)":             g_ColZokusei = CLng(val)
            Case "列番号_機械品番(H列)":         g_ColKikiHinban = CLng(val)
            Case "問い合わせ先メール":            g_InquiryEmail = val
        End Select
    Next i

    Exit Sub
ErrHandler:
    MsgBox "設定シートの読み込みに失敗しました。" & vbCrLf & _
           "設定シートの内容を確認してください。" & vbCrLf & _
           "エラー: " & Err.Description, vbCritical, "設定読み込みエラー"
    End
End Sub
```

- [ ] **Step 2: VBAエディタの「実行」→「Sub/ユーザーフォームの実行」で `設定読み込み` を単独実行し、エラーが出ないことを確認する**

- [ ] **Step 3: イミディエイトウィンドウ（Ctrl+G）で設定値が正しく読み込まれているか確認する**

```vba
? g_ColKPNo    ' → 18 と表示されること
? g_ColModel   ' → 21 と表示されること
? g_TargetSheetName  ' → 日程表 と表示されること
```

---

### Task 3: ModLog — ログ書き込みモジュール

**Files:**
- VBAモジュール `ModLog` を追加

- [ ] **Step 1: 新規モジュール `ModLog` を挿入して以下を貼り付ける**

```vba
Option Explicit

' ログシートに1行書き込む
' result: "成功" / "警告" / "エラー" / "情報"
Public Sub ログ書込(stepName As String, result As String, message As String)
    Dim ws As Worksheet
    Dim nextRow As Long

    ws = ThisWorkbook.Sheets("ログ")
    nextRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1

    ws.Cells(nextRow, 1).Value = Now()
    ws.Cells(nextRow, 2).Value = stepName
    ws.Cells(nextRow, 3).Value = result
    ws.Cells(nextRow, 4).Value = message

    ' 警告・エラーは橙・赤でハイライト
    Select Case result
        Case "警告"
            ws.Cells(nextRow, 3).Interior.Color = RGB(255, 165, 0)
        Case "エラー"
            ws.Cells(nextRow, 3).Interior.Color = RGB(255, 100, 100)
    End Select
End Sub
```

- [ ] **Step 2: イミディエイトウィンドウで動作確認する**

```vba
Call ログ書込("テスト", "情報", "ログ書き込みテスト")
```
期待: ログシートの2行目に記録される

---

### Task 4: ModError — エラーハンドリングモジュール

**Files:**
- VBAモジュール `ModError` を追加

- [ ] **Step 1: 新規モジュール `ModError` を挿入して以下を貼り付ける**

```vba
Option Explicit

' 処理停止エラー: 該当行を黄色ハイライトしてポップアップ表示後に停止
' ws: 加工対象シート, rowNum: 問題のある行番号, message: 表示メッセージ
Public Sub 処理停止エラー(ws As Worksheet, rowNum As Long, message As String)
    ' 該当行を黄色でハイライト
    ws.Rows(rowNum).Interior.Color = RGB(255, 255, 0)

    ' ログに記録
    Call ログ書込("エラー検出", "エラー", "行" & rowNum & ": " & message)

    ' ポップアップ表示
    MsgBox "【処理停止】" & vbCrLf & vbCrLf & _
           message & vbCrLf & vbCrLf & _
           "行番号: " & rowNum & vbCrLf & vbCrLf & _
           "オムロン担当者に問い合わせ後、データを修正して最初から再実行してください。", _
           vbCritical, "生産計画自動化 - 処理停止"

    ' 処理を終了
    End
End Sub

' 警告（続行）: ログに記録してポップアップなしで続行
Public Sub 警告ログ(stepName As String, rowNum As Long, message As String)
    Call ログ書込(stepName, "警告", "行" & rowNum & ": " & message)
End Sub
```

- [ ] **Step 2: イミディエイトウィンドウで `警告ログ` を動作確認する（`処理停止エラー`は実際にEndするため確認のみ）**

```vba
' ログシートに警告行が記録されることを確認
Call 警告ログ("テスト", 5, "テスト警告メッセージ")
```

- [ ] **Step 3: コミット**

```bash
git add 生産計画_マクロ.xlsm
git commit -m "feat: VBA基盤モジュール(ModConfig/ModLog/ModError)を実装"
```

---

## Chunk 3: ステップ⑤〜⑦ 実装

### Task 5: ModStep05 — 計画生産対象削除

**概要:** K列（追加仕様）に「計画生産対象」を含む行を削除する。

**Files:**
- VBAモジュール `ModStep05` を追加

- [ ] **Step 1: テスト用データを「日程表」シートに手動で準備する**

以下の内容で「日程表」シートを加工対象ファイルに用意する（実際のBHプランから出力したもの、またはサンプルデータ）:
- K列に「計画生産対象」が含まれる行: 数行
- K列に別の値や空欄の行: 数行

- [ ] **Step 2: 新規モジュール `ModStep05` を挿入して以下を貼り付ける**

```vba
Option Explicit

' ステップ⑤: K列（追加仕様）に「計画生産対象」を含む行を削除する
Public Sub Step05_計画生産対象削除(ws As Worksheet)
    Dim lastRow As Long
    Dim i As Long
    Dim cellVal As String
    Dim deletedCount As Long

    deletedCount = 0
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    ' 下から上に向かって削除（行削除時のインデックスズレを防ぐ）
    For i = lastRow To 2 Step -1
        cellVal = Trim(ws.Cells(i, g_ColTsuikashiyo).Value)
        If InStr(cellVal, "計画生産対象") > 0 Then
            ws.Rows(i).Delete
            deletedCount = deletedCount + 1
        End If
    Next i

    Call ログ書込("Step05_計画生産対象削除", "成功", deletedCount & "行を削除しました")
End Sub
```

- [ ] **Step 3: テストデータで実行して動作確認する**

イミディエイトウィンドウで:
```vba
Call 設定読み込み()
Dim ws As Worksheet
Set ws = Workbooks("対象ファイル名.xlsx").Sheets("日程表")
Call Step05_計画生産対象削除(ws)
```
期待: 「計画生産対象」を含む行だけが削除され、ログシートに削除件数が記録される

---

### Task 6: ModStep06 — 出荷済みデータ削除

**概要:** N列（光真ss出荷日）が当月より前の行について、R列（KP-No）をBH計画保存版（V8/V9）のJ列と照合し、一致するものを削除する。

**Files:**
- VBAモジュール `ModStep06` を追加

- [ ] **Step 1: 新規モジュール `ModStep06` を挿入して以下を貼り付ける**

```vba
Option Explicit

' ステップ⑥: 出荷済みデータ削除
' BH計画保存版（V8/V9）のKP-Noと照合し、過去月分で一致するものを削除する
Public Sub Step06_出荷済みデータ削除(ws As Worksheet)
    ' BH計画保存版のKP-Noをメモリに読み込む
    Dim savedKPNos As Collection
    Set savedKPNos = 保存版KPNo読み込み()

    Dim lastRow As Long
    Dim i As Long
    Dim kpNo As String
    Dim shukkaDate As Variant
    Dim deletedCount As Long

    deletedCount = 0
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    For i = lastRow To 2 Step -1
        shukkaDate = ws.Cells(i, g_ColShukkaDate).Value

        ' N列が空欄はスキップ
        If shukkaDate = "" Or IsEmpty(shukkaDate) Then GoTo Continue

        ' N列の出荷日が当月より前のもの（過去分）のみチェック対象
        If CDate(shukkaDate) < g_BaseDate Then
            kpNo = Trim(ws.Cells(i, g_ColKPNo).Value)
            If kpNo <> "" Then
                ' 保存版のKP-Noリストに存在すれば出荷済みとして削除
                If KPNoExists(savedKPNos, kpNo) Then
                    ws.Rows(i).Delete
                    deletedCount = deletedCount + 1
                End If
            End If
        End If
Continue:
    Next i

    Call ログ書込("Step06_出荷済みデータ削除", "成功", deletedCount & "行を削除しました")
End Sub

' BH計画保存版（V8/V9）からKP-Noをすべて読み込んでCollectionで返す
Private Function 保存版KPNo読み込み() As Collection
    Dim col As New Collection
    Dim paths(1) As String
    paths(0) = g_V8SavedPath
    paths(1) = g_V9SavedPath

    Dim filePath As String
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim kpNo As String

    ' KP-Noは保存版のJ列（列番号10）に存在する
    ' ※ 実際のファイルで列番号を確認して設定シートに追加すること
    Const COL_KP_SAVED As Long = 10  ' 保存版のKP-No列

    For Each filePath In paths
        If filePath = "" Then GoTo NextFile
        If Dir(filePath) = "" Then
            Call ログ書込("Step06", "警告", "保存版ファイルが見つかりません: " & filePath)
            GoTo NextFile
        End If

        Set wb = Workbooks.Open(filePath, ReadOnly:=True)
        For Each ws In wb.Sheets
            lastRow = ws.Cells(ws.Rows.Count, COL_KP_SAVED).End(xlUp).Row
            For i = 2 To lastRow
                kpNo = Trim(ws.Cells(i, COL_KP_SAVED).Value)
                If kpNo <> "" Then
                    On Error Resume Next
                    col.Add kpNo, kpNo  ' キー重複は無視
                    On Error GoTo 0
                End If
            Next i
        Next ws
        wb.Close SaveChanges:=False
NextFile:
    Next filePath

    Set 保存版KPNo読み込み = col
End Function

' Collection内にkpNoが存在するか確認
Private Function KPNoExists(col As Collection, kpNo As String) As Boolean
    On Error Resume Next
    Dim dummy As String
    dummy = col(kpNo)
    KPNoExists = (Err.Number = 0)
    On Error GoTo 0
End Function
```

- [ ] **Step 2: テスト実行して動作確認する**

確認ポイント:
- 過去月 + 保存版に存在するKP-Noの行が削除される
- 当月以降の行は削除されない
- KP-Noが空欄の行はスキップされる
- ログに削除件数が記録される

---

### Task 7: ModStep07 — 型式（S列）補完

**概要:** S列（BH型式TYPE）が空欄の行に型式を補完する。
- G列が「YB-」または「YU-」で始まる場合: G列の値をそのままS列に設定
- それ以外（3S2Y等）: C列（客先名）が同じ行を検索し、S列に値があるものを取得。G列も同じものを優先
- 3ヶ月以内で補完不可: 処理停止エラー
- 4ヶ月以降で補完不可: 警告ログで続行

**Files:**
- VBAモジュール `ModStep07` を追加

- [ ] **Step 1: 新規モジュール `ModStep07` を挿入して以下を貼り付ける**

```vba
Option Explicit

' ステップ⑦: S列（BH型式TYPE）補完
Public Sub Step07_型式補完(ws As Worksheet)
    Dim lastRow As Long
    Dim i As Long
    Dim bhType As String
    Dim katashiki As String  ' G列
    Dim supplemented As Long
    Dim warned As Long

    supplemented = 0
    warned = 0
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    For i = 2 To lastRow
        bhType = Trim(ws.Cells(i, g_ColBHType).Value)
        If bhType <> "" Then GoTo Continue  ' すでに入力済みはスキップ

        katashiki = Trim(ws.Cells(i, g_ColKatashiki).Value)  ' G列
        Dim shukkaDate As Variant
        shukkaDate = ws.Cells(i, g_ColShukkaDate).Value

        ' 3ヶ月ルール判定
        Dim is3MonthsOrLess As Boolean
        is3MonthsOrLess = False
        If shukkaDate <> "" And Not IsEmpty(shukkaDate) Then
            Dim months3Later As Date
            months3Later = DateSerial(Year(g_BaseDate), Month(g_BaseDate) + 3, Day(g_BaseDate))
            is3MonthsOrLess = (CDate(shukkaDate) <= months3Later)
        End If

        ' G列がYB-またはYU-で始まる場合はそのまま使用
        If Left(katashiki, 3) = "YB-" Or Left(katashiki, 3) = "YU-" Then
            ws.Cells(i, g_ColBHType).Value = katashiki
            supplemented = supplemented + 1
            GoTo Continue
        End If

        ' それ以外: 客先名（C列）で検索して補完
        Dim foundType As String
        foundType = 客先名からBHType取得(ws, i, lastRow)

        If foundType <> "" Then
            ws.Cells(i, g_ColBHType).Value = foundType
            supplemented = supplemented + 1
        ElseIf is3MonthsOrLess Then
            ' 3ヶ月以内で補完不可 → 処理停止
            Call 処理停止エラー(ws, i, _
                "S列（BH型式TYPE）が空欄で自動補完できませんでした。" & vbCrLf & _
                "当月から3ヶ月以内のデータのため、オムロン担当者への問い合わせが必要です。" & vbCrLf & _
                "客先名: " & ws.Cells(i, g_ColKyakusakiName).Value)
        Else
            ' 4ヶ月以降で補完不可 → 警告のみで続行
            Call 警告ログ("Step07_型式補完", i, _
                "S列が空欄で補完不可（4ヶ月以降のデータのため続行）。客先名: " & _
                ws.Cells(i, g_ColKyakusakiName).Value)
            warned = warned + 1
        End If
Continue:
    Next i

    Call ログ書込("Step07_型式補完", "成功", _
        supplemented & "行を補完、" & warned & "行を警告スキップ")
End Sub

' 同一客先名の行からBH型式TYPEを取得する
' 複数候補がある場合はG列（型式）が同じものを優先する
Private Function 客先名からBHType取得(ws As Worksheet, targetRow As Long, lastRow As Long) As String
    Dim targetKyakusaki As String
    Dim targetKatashiki As String
    targetKyakusaki = Trim(ws.Cells(targetRow, g_ColKyakusakiName).Value)
    targetKatashiki = Trim(ws.Cells(targetRow, g_ColKatashiki).Value)

    Dim exactMatch As String  ' G列も一致したもの
    Dim partialMatch As String  ' 客先名だけ一致したもの
    exactMatch = ""
    partialMatch = ""

    Dim i As Long
    For i = 2 To lastRow
        If i = targetRow Then GoTo Continue
        Dim bhType As String
        bhType = Trim(ws.Cells(i, g_ColBHType).Value)
        If bhType = "" Then GoTo Continue

        Dim kyakusaki As String
        kyakusaki = Trim(ws.Cells(i, g_ColKyakusakiName).Value)
        If kyakusaki <> targetKyakusaki Then GoTo Continue

        ' 客先名一致
        If partialMatch = "" Then partialMatch = bhType

        ' G列も一致するものを優先
        Dim katashiki As String
        katashiki = Trim(ws.Cells(i, g_ColKatashiki).Value)
        If katashiki = targetKatashiki Then
            exactMatch = bhType
            Exit For
        End If
Continue:
    Next i

    If exactMatch <> "" Then
        客先名からBHType取得 = exactMatch
    Else
        客先名からBHType取得 = partialMatch
    End If
End Function
```

- [ ] **Step 2: テスト実行して動作確認する**

確認ポイント:
- YB- / YU- 始まりの行はG列の値がそのままS列に入る
- 客先名が同じ行からS列値を参照して補完される
- G列も一致するものが優先される
- 3ヶ月以内で補完不可の場合は処理が停止してポップアップが表示される
- 4ヶ月以降で補完不可の場合はログに警告が記録されて続行する

- [ ] **Step 3: コミット**

```bash
git add 生産計画_マクロ.xlsm
git commit -m "feat: Step05(計画生産対象削除)/Step06(出荷済み削除)/Step07(型式補完)を実装"
```

---

## Chunk 4: ステップ⑧〜⑩ 実装

### Task 8: ModStep08 — 計画生産行展開（1台1行化）

**概要:** F列（機種名）に「計画生産」を含み、かつN列（出荷日）が3ヶ月以内の行を対象に、L列（数量）の数だけ行をコピーして展開する。B列（生産計画No）の末尾に -01, -02... の連番を付与する。

**Files:**
- VBAモジュール `ModStep08` を追加

- [ ] **Step 1: 新規モジュール `ModStep08` を挿入して以下を貼り付ける**

```vba
Option Explicit

' ステップ⑧: 計画生産行展開（1台1行化）
' 対象: F列に「計画生産」を含む + N列が当月〜3ヶ月以内
Public Sub Step08_計画生産行展開(ws As Worksheet)
    Dim lastRow As Long
    Dim i As Long
    Dim kishuName As String
    Dim suryo As Long
    Dim shukkaDate As Variant
    Dim months3Later As Date
    Dim expandedCount As Long

    expandedCount = 0
    months3Later = DateSerial(Year(g_BaseDate), Month(g_BaseDate) + 3, Day(g_BaseDate))

    ' 下から処理すると行挿入後のインデックスが安定する
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    For i = lastRow To 2 Step -1
        kishuName = Trim(ws.Cells(i, g_ColKishuName).Value)
        If InStr(kishuName, "計画生産") = 0 Then GoTo Continue

        shukkaDate = ws.Cells(i, g_ColShukkaDate).Value
        If shukkaDate = "" Or IsEmpty(shukkaDate) Then GoTo Continue
        If CDate(shukkaDate) > months3Later Then GoTo Continue

        suryo = CLng(ws.Cells(i, g_ColSuryo).Value)
        If suryo <= 1 Then GoTo Continue  ' 数量1は展開不要

        ' 元の生産計画Noを取得
        Dim baseNo As String
        baseNo = Trim(ws.Cells(i, g_ColSeisanNo).Value)

        ' 数量分の行をi+1以降に挿入してコピー
        Dim j As Long
        For j = suryo To 1 Step -1
            If j > 1 Then
                ws.Rows(i + 1).Insert Shift:=xlDown
                ws.Rows(i).Copy ws.Rows(i + 1)
            End If
            ' 連番付与: baseNo-01, -02 ...
            ws.Cells(i + j - 1, g_ColSeisanNo).Value = baseNo & "-" & Format(j, "00")
            ' 数量を1に更新
            ws.Cells(i + j - 1, g_ColSuryo).Value = 1
        Next j

        expandedCount = expandedCount + 1
Exit For  ' ← この行削除（下記に正しいContinueラベルあり）
Continue:
    Next i

    Call ログ書込("Step08_計画生産行展開", "成功", expandedCount & "件の行展開を実施しました")
End Sub
```

> **注意:** 上記コードの `Exit For` の行は誤りのため削除してください。`Continue:` ラベルの直前に `Exit For` は不要です。

実際に貼り付けるコードは以下の通りです（修正版）:

```vba
Option Explicit

Public Sub Step08_計画生産行展開(ws As Worksheet)
    Dim months3Later As Date
    Dim expandedCount As Long
    expandedCount = 0
    months3Later = DateSerial(Year(g_BaseDate), Month(g_BaseDate) + 3, Day(g_BaseDate))

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    Dim i As Long
    For i = lastRow To 2 Step -1
        Dim kishuName As String
        kishuName = Trim(ws.Cells(i, g_ColKishuName).Value)
        If InStr(kishuName, "計画生産") = 0 Then GoTo NextRow

        Dim shukkaDate As Variant
        shukkaDate = ws.Cells(i, g_ColShukkaDate).Value
        If IsEmpty(shukkaDate) Or shukkaDate = "" Then GoTo NextRow
        If CDate(shukkaDate) > months3Later Then GoTo NextRow

        Dim suryo As Long
        suryo = CLng(ws.Cells(i, g_ColSuryo).Value)
        If suryo <= 1 Then GoTo NextRow

        Dim baseNo As String
        baseNo = Trim(ws.Cells(i, g_ColSeisanNo).Value)

        Dim j As Long
        For j = suryo To 1 Step -1
            If j > 1 Then
                ws.Rows(i + 1).Insert Shift:=xlDown
                ws.Rows(i).Copy ws.Rows(i + 1)
            End If
            ws.Cells(i + j - 1, g_ColSeisanNo).Value = baseNo & "-" & Format(j, "00")
            ws.Cells(i + j - 1, g_ColSuryo).Value = 1
        Next j

        expandedCount = expandedCount + 1
NextRow:
    Next i

    Call ログ書込("Step08_計画生産行展開", "成功", expandedCount & "件の行展開を実施しました")
End Sub
```

- [ ] **Step 2: テスト実行して動作確認する**

確認ポイント:
- 数量6の計画生産行が6行に展開される
- 生産計画Noの末尾に -01〜-06 の連番が付く
- 各行の数量が1になる
- 4ヶ月以降の計画生産行は対象外で展開されない

---

### Task 9: ModStep09 — 数量チェック

**概要:** MODEL（U列）が「V8」または「V9」（メンテ以外）の行で、N列（出荷日）が3ヶ月以内のものに数量1以外があればエラー停止する。

**Files:**
- VBAモジュール `ModStep09` を追加

- [ ] **Step 1: 新規モジュール `ModStep09` を挿入して以下を貼り付ける**

```vba
Option Explicit

' ステップ⑨: 数量チェック（V8/V9 3ヶ月以内）
Public Sub Step09_数量チェック(ws As Worksheet)
    Dim lastRow As Long
    Dim i As Long
    Dim model As String
    Dim suryo As Long
    Dim shukkaDate As Variant
    Dim months3Later As Date

    months3Later = DateSerial(Year(g_BaseDate), Month(g_BaseDate) + 3, Day(g_BaseDate))
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    For i = 2 To lastRow
        model = Trim(ws.Cells(i, g_ColModel).Value)

        ' V8またはV9（メンテ除く）のみチェック
        If model <> "V8" And model <> "V9" Then GoTo NextRow

        shukkaDate = ws.Cells(i, g_ColShukkaDate).Value
        If IsEmpty(shukkaDate) Or shukkaDate = "" Then GoTo NextRow
        If CDate(shukkaDate) > months3Later Then GoTo NextRow

        suryo = CLng(ws.Cells(i, g_ColSuryo).Value)
        If suryo <> 1 Then
            Call 処理停止エラー(ws, i, _
                "MODEL「" & model & "」で数量が1ではありません（数量=" & suryo & "）。" & vbCrLf & _
                "オムロン担当者への問い合わせが必要です。" & vbCrLf & _
                "生産計画No: " & ws.Cells(i, g_ColSeisanNo).Value)
        End If
NextRow:
    Next i

    Call ログ書込("Step09_数量チェック", "成功", "V8/V9の3ヶ月以内数量チェック完了")
End Sub
```

- [ ] **Step 2: テスト実行して動作確認する**

確認ポイント:
- V8/V9で数量>1かつ3ヶ月以内の行があれば処理が停止する
- メンテV8/メンテV9はチェック対象外
- 4ヶ月以降はチェック対象外

---

### Task 10: ModStep10 — 並び替え

**概要:**
- V8/V9: MODEL → 光真ss出荷日 → 順序指示発行日 → KP-No → 属性（降順）→ 客先名 → 生産計画No
- メンテV8/メンテV9: MODEL → 機械品番 → 光真ss出荷日 → 順序指示発行日 → KP-No

**Files:**
- VBAモジュール `ModStep10` を追加

- [ ] **Step 1: 新規モジュール `ModStep10` を挿入して以下を貼り付ける**

```vba
Option Explicit

' ステップ⑩: 並び替え
Public Sub Step10_並び替え(ws As Worksheet)
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastRow < 2 Then
        Call ログ書込("Step10_並び替え", "情報", "データなし、スキップ")
        Exit Sub
    End If

    Dim sortRange As Range
    Set sortRange = ws.Range(ws.Cells(2, 1), ws.Cells(lastRow, ws.UsedRange.Columns.Count))

    With ws.Sort
        .SortFields.Clear

        ' 1. MODEL(U列): 昇順
        .SortFields.Add Key:=ws.Columns(g_ColModel), Order:=xlAscending
        ' 2. 機械品番(H列): 昇順（メンテ系の2番目キー）
        .SortFields.Add Key:=ws.Columns(g_ColKikiHinban), Order:=xlAscending
        ' 3. 光真ss出荷日(N列): 昇順
        .SortFields.Add Key:=ws.Columns(g_ColShukkaDate), Order:=xlAscending
        ' 4. 順序指示発行日(M列): 昇順
        .SortFields.Add Key:=ws.Columns(g_ColJunjoHakkoDate), Order:=xlAscending
        ' 5. KP-No(R列): 昇順
        .SortFields.Add Key:=ws.Columns(g_ColKPNo), Order:=xlAscending
        ' 6. 属性(I列): 降順
        .SortFields.Add Key:=ws.Columns(g_ColZokusei), Order:=xlDescending
        ' 7. 客先名(C列): 昇順
        .SortFields.Add Key:=ws.Columns(g_ColKyakusakiName), Order:=xlAscending
        ' 8. 生産計画No(B列): 昇順
        .SortFields.Add Key:=ws.Columns(g_ColSeisanNo), Order:=xlAscending

        .SetRange sortRange
        .Header = xlNo
        .Apply
    End With

    Call ログ書込("Step10_並び替え", "成功", "並び替え完了（" & lastRow - 1 & "行）")
End Sub
```

> **補足:** VBAのSort.SortFieldsは最大64キーまで登録できるが、実際の適用優先順位は登録順。MODELを最初に登録することでV8/V9/メンテの塊ごとに並ぶ。その中でメンテ系は機械品番が有意なキーになるため2番目に設定している。

- [ ] **Step 2: テスト実行して動作確認する**

確認ポイント:
- V8/V9のグループとメンテV8/V9のグループが分かれて並ぶ
- 同一MODEL内は出荷日昇順に並ぶ

- [ ] **Step 3: コミット**

```bash
git add 生産計画_マクロ.xlsm
git commit -m "feat: Step08(行展開)/Step09(数量チェック)/Step10(並び替え)を実装"
```

---

## Chunk 5: メイン実行とUI

### Task 11: ModMain — メイン実行と実行ボタン

**Files:**
- VBAモジュール `ModMain` を追加
- 「設定」シートに「実行」ボタンを追加

- [ ] **Step 1: 新規モジュール `ModMain` を挿入して以下を貼り付ける**

```vba
Option Explicit

' メイン実行: Phase 1-A（ステップ⑤〜⑩）
Public Sub メイン実行()
    ' 開始確認
    Dim ans As VbMsgBoxResult
    ans = MsgBox("生産計画自動化を開始します。" & vbCrLf & vbCrLf & _
                 "加工対象ファイルが正しく配置されていることを確認してください。" & vbCrLf & _
                 "（BHプランの出力ファイルを input フォルダに置いてください）" & vbCrLf & vbCrLf & _
                 "続行しますか？", vbYesNo + vbQuestion, "生産計画自動化")
    If ans = vbNo Then Exit Sub

    ' 設定読み込み
    Call 設定読み込み()

    ' 加工対象ファイルを開く
    Dim targetWb As Workbook
    Dim targetWs As Worksheet
    Set targetWb = 対象ファイルを開く()
    If targetWb Is Nothing Then Exit Sub
    Set targetWs = targetWb.Sheets(g_TargetSheetName)

    ' ログに開始を記録
    Call ログ書込("メイン実行", "情報", "処理開始: " & targetWb.Name)

    ' ===== ステップ実行 =====
    Call Step05_計画生産対象削除(targetWs)
    Call Step06_出荷済みデータ削除(targetWs)
    Call Step07_型式補完(targetWs)
    Call Step08_計画生産行展開(targetWs)
    Call Step09_数量チェック(targetWs)
    Call Step10_並び替え(targetWs)
    ' ========================

    ' 上書き保存
    targetWb.Save

    Call ログ書込("メイン実行", "成功", "Phase 1-A 処理完了")
    MsgBox "処理が完了しました。" & vbCrLf & _
           "ログシートで処理結果を確認してください。", vbInformation, "完了"
End Sub

' input フォルダ内の最新xlsxファイルを開く
Private Function 対象ファイルを開く() As Workbook
    Dim folderPath As String
    folderPath = g_BHPlanFolder
    If Right(folderPath, 1) <> "\" Then folderPath = folderPath & "\"

    Dim fileName As String
    fileName = Dir(folderPath & "*.xlsx")

    If fileName = "" Then
        MsgBox "inputフォルダにxlsxファイルが見つかりません。" & vbCrLf & _
               "フォルダ: " & folderPath, vbCritical, "ファイルなし"
        Set 対象ファイルを開く = Nothing
        Exit Function
    End If

    ' 複数ある場合は最新のものを使う（簡易: 最初に見つかったもの）
    ' TODO: 必要に応じて日付順でソートして最新を選択する
    Set 対象ファイルを開く = Workbooks.Open(folderPath & fileName)
End Function
```

- [ ] **Step 2: 「設定」シートにボタンを追加してマクロを割り当てる**

VBAエディタではなくExcel画面で:
1. 「開発」タブ → 「挿入」→「ボタン（フォームコントロール）」
2. シート上にボタンを描画
3. マクロの割り当てで `メイン実行` を選択
4. ボタンのテキストを「▶ 生産計画自動化 実行」に変更

- [ ] **Step 3: ボタンを押して一連の処理が流れることをエンドツーエンドで確認する**

テストシナリオ:
1. テスト用日程表データを input フォルダに配置
2. 「▶ 生産計画自動化 実行」ボタンをクリック
3. ⑤〜⑩の各ステップが順番に実行される
4. ログシートに各ステップの結果が記録される
5. 加工済みファイルが保存される

- [ ] **Step 4: 最終コミット**

```bash
git add 生産計画_マクロ.xlsm
git commit -m "feat: ModMain(メイン実行・実行ボタン)を実装しPhase 1-A完成"
```

---

## 次のフェーズ

Phase 1-A 完了後:
- **Phase 1-B**: ステップ⑪〜㉖（星取表・集計・KMP計画） → 別計画書
- **Phase 1-C**: UiPath フロー（VBA呼び出し一本化） → Phase 1-B完了後に作成

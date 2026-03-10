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

Public g_InquiryEmail           As String  ' 問い合わせ先メール

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

    ' A列=キー, B列=値 の形式で2行目から読み込む
    For i = 2 To ws.UsedRange.Rows.Count + 1
        key = Trim(CStr(ws.Cells(i, 1).Value))
        val = Trim(CStr(ws.Cells(i, 2).Value))
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
        End Select
    Next i

    Exit Sub
ErrHandler:
    MsgBox "設定シートの読み込みに失敗しました。" & vbCrLf & _
           "設定シートの内容を確認してください。" & vbCrLf & _
           "エラー: " & Err.Description, vbCritical, "設定読み込みエラー"
    End
End Sub

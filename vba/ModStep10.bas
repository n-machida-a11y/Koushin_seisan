Attribute VB_Name = "ModStep10"
Option Explicit

' ============================================================
' ステップ⑩: 並び替え
'
' 全モデル共通の優先順位（上位キーから）:
'   1. MODEL(U列): 昇順
'   2. 機械品番(H列): 昇順 ※メンテ系の第2キー、V8/V9は空欄なので実質スキップ
'   3. 光真ss出荷日(N列): 昇順
'   4. 順序指示発行日(M列): 昇順
'   5. KP-No(R列): 昇順
'   6. 属性(I列): 降順
'   7. 客先名(C列): 昇順
'   8. 生産計画No(B列): 昇順
'
' ※ V8/V9の機械品番は空欄のため、実質 V8/V9とメンテをMODELで分けた後、
'   V8/V9は出荷日→発行日→KP-No→属性→客先名→計画Noの順になる
' ============================================================
Public Sub Step10_並び替え(ws As Worksheet)
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    If lastRow < 2 Then
        Call ログ書込("Step10_並び替え", "情報", "データなし、スキップ")
        Exit Sub
    End If

    Dim lastCol As Long
    lastCol = ws.UsedRange.Columns.Count

    Dim sortRange As Range
    Set sortRange = ws.Range(ws.Cells(2, 1), ws.Cells(lastRow, lastCol))

    With ws.Sort
        .SortFields.Clear
        .SortFields.Add Key:=ws.Columns(g_ColModel),         Order:=xlAscending   ' 1. MODEL
        .SortFields.Add Key:=ws.Columns(g_ColKikiHinban),    Order:=xlAscending   ' 2. 機械品番
        .SortFields.Add Key:=ws.Columns(g_ColShukkaDate),    Order:=xlAscending   ' 3. 出荷日
        .SortFields.Add Key:=ws.Columns(g_ColJunjoHakkoDate),Order:=xlAscending   ' 4. 発行日
        .SortFields.Add Key:=ws.Columns(g_ColKPNo),          Order:=xlAscending   ' 5. KP-No
        .SortFields.Add Key:=ws.Columns(g_ColZokusei),       Order:=xlDescending  ' 6. 属性（降順）
        .SortFields.Add Key:=ws.Columns(g_ColKyakusakiName), Order:=xlAscending   ' 7. 客先名
        .SortFields.Add Key:=ws.Columns(g_ColSeisanNo),      Order:=xlAscending   ' 8. 生産計画No
        .SetRange sortRange
        .Header = xlNo
        .Apply
    End With

    Call ログ書込("Step10_並び替え", "成功", "並び替え完了（" & lastRow - 1 & "行）")
End Sub

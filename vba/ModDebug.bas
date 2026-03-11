Attribute VB_Name = "ModDebug"
Option Explicit

' ============================================================
' 診断用モジュール（デバッグ後は削除してください）
' ============================================================
Public Sub KPNo照合診断()
    On Error GoTo ErrHandler
    Call 設定読み込み()

    Dim msg As String
    msg = "=== KP-No 照合診断 ===" & vbCrLf & vbCrLf

    ' --- 保存版のKP-Noサンプル ---
    msg = msg & "[保存版V8] パス: " & g_V8SavedPath & vbCrLf
    msg = msg & "[保存版V8] KPNo列番号: " & g_V8SavedKPNoCol & vbCrLf
    msg = msg & 保存版サンプル(g_V8SavedPath, g_V8SavedKPNoCol) & vbCrLf

    msg = msg & "[保存版V9] パス: " & g_V9SavedPath & vbCrLf
    msg = msg & "[保存版V9] KPNo列番号: " & g_V9SavedKPNoCol & vbCrLf
    msg = msg & 保存版サンプル(g_V9SavedPath, g_V9SavedKPNoCol) & vbCrLf

    ' --- 加工対象のKP-Noサンプル ---
    msg = msg & "[加工対象] KPNo列番号(g_ColKPNo): " & g_ColKPNo & vbCrLf
    msg = msg & "[加工対象] 出荷日列番号(g_ColShukkaDate): " & g_ColShukkaDate & vbCrLf
    msg = msg & 対象サンプル() & vbCrLf

    MsgBox msg, vbInformation, "KP-No 照合診断"
    Exit Sub
ErrHandler:
    MsgBox "エラー発生" & vbCrLf & _
           "エラー番号: " & Err.Number & vbCrLf & _
           "エラー内容: " & Err.Description, vbCritical, "診断エラー"
End Sub

' 保存版から先頭5件をそのまま読む（型変換なし）
Private Function 保存版サンプル(filePath As String, kpNoCol As Long) As String
    If filePath = "" Then
        保存版サンプル = "  → パス未設定" & vbCrLf
        Exit Function
    End If

    Dim exists As Boolean
    On Error Resume Next
    exists = (Dir(filePath) <> "")
    On Error GoTo 0
    If Not exists Then
        保存版サンプル = "  → ファイルなし" & vbCrLf
        Exit Function
    End If

    Dim wb As Workbook
    Set wb = Workbooks.Open(filePath, ReadOnly:=True)

    Dim result As String
    result = "  シート一覧: "
    Dim ws As Worksheet
    For Each ws In wb.Sheets
        result = result & ws.Name & " / "
    Next ws
    result = result & vbCrLf

    Dim found As Long
    For Each ws In wb.Sheets
        Dim r As Long
        For r = 1 To 10
            Dim v As Variant
            v = ws.Cells(r, kpNoCol).Value
            result = result & "  [" & ws.Name & "]" & r & "行目: 値=" & CStr(v) & " 型=" & TypeName(v) & vbCrLf
            found = found + 1
            If found >= 5 Then Exit For
        Next r
        If found >= 5 Then Exit For
    Next ws

    wb.Close SaveChanges:=False
    保存版サンプル = result
End Function

' 加工対象から先頭5件をそのまま読む（型変換なし）
Private Function 対象サンプル() As String
    Dim folderPath As String
    folderPath = g_BHPlanFolder
    If Right(folderPath, 1) <> "\" Then folderPath = folderPath & "\"

    Dim fileName As String
    Dim latestFile As String
    Dim latestDate As Date
    On Error Resume Next
    fileName = Dir(folderPath & "*.xlsx")
    On Error GoTo 0

    Do While fileName <> ""
        Dim d As Date
        d = FileDateTime(folderPath & fileName)
        If d > latestDate Then
            latestDate = d
            latestFile = fileName
        End If
        fileName = Dir()
    Loop

    If latestFile = "" Then
        対象サンプル = "  → inputフォルダにxlsxなし" & vbCrLf
        Exit Function
    End If

    Dim wb As Workbook
    Set wb = Workbooks.Open(folderPath & latestFile, ReadOnly:=True)

    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Sheets(g_TargetSheetName)
    On Error GoTo 0
    If ws Is Nothing Then
        対象サンプル = "  → シート[" & g_TargetSheetName & "]なし" & vbCrLf
        wb.Close SaveChanges:=False
        Exit Function
    End If

    Dim result As String
    result = "  ファイル: " & latestFile & vbCrLf
    Dim found As Long
    Dim i As Long
    For i = 2 To ws.UsedRange.Rows.Count + 1
        ' R列(KPNo)とN列(出荷日)をそのまま表示（変換なし）
        Dim kpRaw As Variant
        Dim dtRaw As Variant
        kpRaw = ws.Cells(i, g_ColKPNo).Value
        dtRaw = ws.Cells(i, g_ColShukkaDate).Value
        If Not IsEmpty(kpRaw) And CStr(kpRaw) <> "" Then
            result = result & "  行" & i & ": KP値=" & CStr(kpRaw) & " KP型=" & TypeName(kpRaw) & _
                     " / 出荷日値=" & CStr(dtRaw) & " 出荷日型=" & TypeName(dtRaw) & vbCrLf
            found = found + 1
            If found >= 5 Then Exit For
        End If
    Next i

    wb.Close SaveChanges:=False
    対象サンプル = result
End Function

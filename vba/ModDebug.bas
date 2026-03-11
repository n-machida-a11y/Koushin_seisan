Attribute VB_Name = "ModDebug"
Option Explicit

' ============================================================
' 診断用モジュール（デバッグ後は削除してください）
' Step06 KP-No照合の不一致を調査するためのツール
' ============================================================

' 実行方法: VBAエディタのイミディエイトウィンドウで
'   Call KPNo照合診断()
' ============================================================
Public Sub KPNo照合診断()
    Call 設定読み込み()

    Dim msg As String
    msg = "=== KP-No 照合診断 ===" & vbCrLf & vbCrLf

    ' --- (1) 保存版から読み込んだKP-Noのサンプルを表示 ---
    msg = msg & "【1】保存版から読み込んだKP-No（先頭5件）" & vbCrLf
    msg = msg & 保存版KPNoサンプル取得(g_V8SavedPath, g_V8SavedKPNoCol, "V8保存版") & vbCrLf
    msg = msg & 保存版KPNoサンプル取得(g_V9SavedPath, g_V9SavedKPNoCol, "V9保存版") & vbCrLf

    ' --- (2) 加工対象ファイルの過去月データのKP-Noサンプルを表示 ---
    msg = msg & "【2】加工対象データの過去月KP-No（先頭5件）" & vbCrLf
    msg = msg & 対象KPNoサンプル取得() & vbCrLf

    MsgBox msg, vbInformation, "KP-No 照合診断"
End Sub

Private Function 保存版KPNoサンプル取得(filePath As String, kpNoCol As Long, label As String) As String
    Dim result As String
    result = label & ": "

    If filePath = "" Then
        保存版KPNoサンプル取得 = result & "パス未設定" & vbCrLf
        Exit Function
    End If

    Dim fileExists As Boolean
    Dim dirErr As Long
    On Error Resume Next
    fileExists = (Dir(filePath) <> "")
    dirErr = Err.Number
    On Error GoTo 0

    If dirErr <> 0 Or Not fileExists Then
        保存版KPNoサンプル取得 = result & "ファイルなし(" & filePath & ")" & vbCrLf
        Exit Function
    End If

    Dim wb As Workbook
    Set wb = Workbooks.Open(filePath, ReadOnly:=True)

    result = result & vbCrLf
    result = result & "  列番号=" & kpNoCol & "(" & シート名一覧(wb) & ")" & vbCrLf

    Dim ws As Worksheet
    Dim found As Long
    found = 0
    For Each ws In wb.Sheets
        Dim lastRow As Long
        lastRow = ws.Cells(ws.Rows.Count, kpNoCol).End(xlUp).Row
        Dim i As Long
        For i = 2 To lastRow
            If found >= 5 Then Exit For
            Dim v As Variant
            v = ws.Cells(i, kpNoCol).Value
            If Not IsEmpty(v) And CStr(v) <> "" Then
                result = result & "  [" & ws.Name & "]行" & i & ": " & _
                         "値=" & CStr(v) & " 型=" & TypeName(v) & vbCrLf
                found = found + 1
            End If
        Next i
        If found >= 5 Then Exit For
    Next ws

    If found = 0 Then
        result = result & "  → 列" & kpNoCol & "にデータなし。列番号を確認してください" & vbCrLf
    End If

    wb.Close SaveChanges:=False
    保存版KPNoサンプル取得 = result
End Function

Private Function 対象KPNoサンプル取得() As String
    Dim result As String
    result = ""

    ' inputフォルダの最新xlsxを開く
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
        Dim fileDate As Date
        fileDate = FileDateTime(folderPath & fileName)
        If fileDate > latestDate Then
            latestDate = fileDate
            latestFile = fileName
        End If
        fileName = Dir()
    Loop

    If latestFile = "" Then
        対象KPNoサンプル取得 = "  inputフォルダにxlsxなし" & vbCrLf
        Exit Function
    End If

    Dim wb As Workbook
    Set wb = Workbooks.Open(folderPath & latestFile, ReadOnly:=True)

    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Sheets(g_TargetSheetName)
    On Error GoTo 0

    If ws Is Nothing Then
        対象KPNoサンプル取得 = "  シート「" & g_TargetSheetName & "」なし" & vbCrLf
        wb.Close SaveChanges:=False
        Exit Function
    End If

    Dim found As Long
    found = 0
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    Dim i As Long
    For i = 2 To lastRow
        If found >= 5 Then Exit For
        Dim shukkaDate As Variant
        shukkaDate = ws.Cells(i, g_ColShukkaDate).Value
        If IsEmpty(shukkaDate) Or CStr(shukkaDate) = "" Then GoTo NextRow

        If CDate(shukkaDate) < g_BaseDate Then
            Dim kpNo As String
            kpNo = Trim(CStr(ws.Cells(i, g_ColKPNo).Value))
            If kpNo <> "" Then
                Dim rawVal As Variant
                rawVal = ws.Cells(i, g_ColKPNo).Value
                result = result & "  行" & i & ": 値=" & kpNo & _
                         " 型=" & TypeName(rawVal) & _
                         " 出荷日=" & CStr(shukkaDate) & vbCrLf
                found = found + 1
            End If
        End If
NextRow:
    Next i

    If found = 0 Then
        result = result & "  過去月かつKP-Noありの行なし（列番号=" & g_ColKPNo & "）" & vbCrLf
    End If

    wb.Close SaveChanges:=False
    対象KPNoサンプル取得 = result
End Function

Private Function シート名一覧(wb As Workbook) As String
    Dim names As String
    Dim ws As Worksheet
    For Each ws In wb.Sheets
        If names <> "" Then names = names & ", "
        names = names & ws.Name
    Next ws
    シート名一覧 = names
End Function

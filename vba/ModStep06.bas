Attribute VB_Name = "ModStep06"
Option Explicit

' ============================================================
' ステップ⑥: 出荷済みデータ削除
' N列（光真ss出荷日）が当月より前の行で、R列（KP-No）が
' BH計画保存版（V8/V9）に存在するものを出荷済みとして削除する
' ============================================================
Public Sub Step06_出荷済みデータ削除(ws As Worksheet)
    Dim savedKPNos As Collection
    Set savedKPNos = 保存版KPNo読み込み()

    Dim lastRow As Long
    Dim i As Long
    Dim kpNo As String
    Dim shukkaDate As Variant
    Dim deletedCount As Long

    deletedCount = 0
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    ' 下から上に向かって処理（行削除時のインデックスズレを防ぐ）
    For i = lastRow To 2 Step -1
        shukkaDate = ws.Cells(i, g_ColShukkaDate).Value

        If IsEmpty(shukkaDate) Or CStr(shukkaDate) = "" Then GoTo NextRow

        ' 出荷日が当月より前のもの（過去分）のみチェック対象
        If CDate(shukkaDate) < g_BaseDate Then
            kpNo = KPNoを正規化(ws.Cells(i, g_ColKPNo).Value)
            If kpNo <> "" Then
                If KPNoExists(savedKPNos, kpNo) Then
                    ws.Rows(i).Delete
                    deletedCount = deletedCount + 1
                End If
            End If
        End If
NextRow:
    Next i

    Call ログ書込("Step06_出荷済みデータ削除", "成功", deletedCount & "行を削除しました")
End Sub

' ============================================================
' BH計画保存版（V8/V9）からKP-Noをすべて読み込んでCollectionで返す
' ============================================================
Private Function 保存版KPNo読み込み() As Collection
    Dim col As New Collection

    Dim pathInfo(1, 1) As Variant
    pathInfo(0, 0) = g_V8SavedPath
    pathInfo(0, 1) = g_V8SavedKPNoCol
    pathInfo(1, 0) = g_V9SavedPath
    pathInfo(1, 1) = g_V9SavedKPNoCol

    Dim idx As Long
    Dim filePath As String
    Dim kpNoCol As Long
    Dim fileExists As Boolean
    Dim dirErrNum As Long
    For idx = 0 To 1
        filePath = CStr(pathInfo(idx, 0))
        kpNoCol = CLng(pathInfo(idx, 1))

        If filePath = "" Then GoTo NextFile
        ' Dir()はドライブが存在しない場合にエラー52を発生させるためOn Error で保護する
        ' Err.NumberはOn Error GoTo 0でクリアされるため、ブロック内で先に取得する
        fileExists = False
        dirErrNum = 0
        On Error Resume Next
        fileExists = (Dir(filePath) <> "")
        dirErrNum = Err.Number
        On Error GoTo 0
        If dirErrNum <> 0 And dirErrNum <> 52 Then
            Call ログ書込("Step06", "警告", "ファイル確認中にエラーが発生しました(Error " & dirErrNum & "): " & filePath)
        End If
        If Not fileExists Then
            Call ログ書込("Step06", "警告", "保存版ファイルが見つかりません: " & filePath)
            GoTo NextFile
        End If

        Dim wb As Workbook
        Set wb = Workbooks.Open(filePath, ReadOnly:=True)

        Dim ws As Worksheet
        For Each ws In wb.Sheets
            Dim lastRow As Long
            lastRow = ws.Cells(ws.Rows.Count, kpNoCol).End(xlUp).Row
            Dim i As Long
            For i = 2 To lastRow
                Dim kpNo As String
                kpNo = KPNoを正規化(ws.Cells(i, kpNoCol).Value)
                If kpNo <> "" Then
                    On Error Resume Next
                    col.Add kpNo, kpNo
                    On Error GoTo 0
                End If
            Next i
        Next ws

        wb.Close SaveChanges:=False
NextFile:
    Next idx

    Set 保存版KPNo読み込み = col
End Function

' ============================================================
' KP-No値を型に関わらず統一した文字列に正規化する
' 数値型（Double/Long等）は小数点・カンマなしの整数文字列に変換
' 文字列型はTrimのみ
' ============================================================
Private Function KPNoを正規化(v As Variant) As String
    If IsEmpty(v) Or IsNull(v) Then
        KPNoを正規化 = ""
    ElseIf IsNumeric(v) Then
        ' 数値型: 整数部のみの文字列に変換（"12345.0" → "12345"）
        KPNoを正規化 = CStr(CLng(v))
    Else
        KPNoを正規化 = Trim(CStr(v))
    End If
End Function

' ============================================================
' Collection内にkpNoが存在するか確認
' ============================================================
Private Function KPNoExists(col As Collection, kpNo As String) As Boolean
    On Error Resume Next
    Dim dummy As String
    dummy = col(kpNo)
    KPNoExists = (Err.Number = 0)
    On Error GoTo 0
End Function

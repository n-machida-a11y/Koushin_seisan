Attribute VB_Name = "ModStep11"
Option Explicit

' ============================================================
' Step11: メンテ行の星取表照合
'
' MODEL=メンテV8/メンテV9の行について、Production Scheduleの
' 星取表シートに同じデータがあるか照合する。
'
' 照合キー: KP-No(R列) or 生産計画No(B列)
' 比較項目: 型式(G列), 順序指示発行日(M列), 光真ss出荷日(N列), 数量(L列)
'
' 判定:
'   全一致 → OK
'   一部一致 → 黄色ハイライト+ログ（全件検出後にまとめて停止）
'   見つからない → 新規 → 星取表に追加
' ============================================================
Public Sub Step11_メンテ照合(targetWs As Worksheet)
    Dim lastRow As Long
    lastRow = targetWs.Cells(targetWs.Rows.Count, 1).End(xlUp).Row
    
    Dim partialCount As Long
    Dim newCount As Long
    Dim matchCount As Long
    partialCount = 0
    newCount = 0
    matchCount = 0
    
    ' V8星取表を開く
    Dim v8HsWs As Worksheet
    Dim v8Wb As Workbook
    If g_V8ProdSchedulePath <> "" Then
        Set v8Wb = Workbooks.Open(g_V8ProdSchedulePath)
        Set v8HsWs = シート検索(v8Wb, g_SheetV8Hoshitori)
    End If
    
    ' V9星取表を開く
    Dim v9HsWs As Worksheet
    Dim v9Wb As Workbook
    If g_V9ProdSchedulePath <> "" Then
        ' V8と同じファイルでなければ開く
        If g_V9ProdSchedulePath <> g_V8ProdSchedulePath Then
            Set v9Wb = Workbooks.Open(g_V9ProdSchedulePath)
        Else
            Set v9Wb = v8Wb
        End If
        Set v9HsWs = シート検索(v9Wb, g_SheetV9Hoshitori)
    End If
    
    ' メンテ行を走査
    Dim i As Long
    For i = g_DataStartRow To lastRow
        Dim model As String
        model = Trim(CStr(targetWs.Cells(i, g_ColModel).Value))
        
        ' メンテV8/V9のみ対象
        If InStr(model, "メンテ") = 0 Then GoTo NextRow
        
        ' 照合先の星取表を選択
        Dim hsWs As Worksheet
        If InStr(model, "V8") > 0 Then
            Set hsWs = v8HsWs
        Else
            Set hsWs = v9HsWs
        End If
        
        If hsWs Is Nothing Then
            Call 警告ログ("Step11_メンテ照合", i, "星取表シートが見つかりません（" & model & "）")
            GoTo NextRow
        End If
        
        ' BHPlanの照合データ
        Dim kpNo As String
        Dim seisanNo As String
        Dim katashiki As String
        Dim junjoDate As Variant
        Dim shukkaDate As Variant
        Dim suryo As Variant
        
        kpNo = Trim(CStr(targetWs.Cells(i, g_ColKPNo).Value))
        seisanNo = Trim(CStr(targetWs.Cells(i, g_ColSeisanNo).Value))
        katashiki = Trim(CStr(targetWs.Cells(i, g_ColKatashiki).Value))
        junjoDate = targetWs.Cells(i, g_ColJunjoHakkoDate).Value
        shukkaDate = targetWs.Cells(i, g_ColShukkaDate).Value
        suryo = targetWs.Cells(i, g_ColSuryo).Value
        
        ' 星取表を検索（KP-Noで検索、なければ生産計画Noで検索）
        Dim hsLastRow As Long
        hsLastRow = hsWs.Cells(hsWs.Rows.Count, 1).End(xlUp).Row
        
        ' 星取表のKP-No列を特定（V8=col13, V9=col10 が一般的だが設定値を使用）
        Dim hsKPCol As Long
        If InStr(model, "V8") > 0 Then
            hsKPCol = g_V8SavedKPNoCol
        Else
            hsKPCol = g_V9SavedKPNoCol
        End If
        
        Dim foundRow As Long
        foundRow = 0
        Dim foundByKP As Boolean
        foundByKP = False
        
        Dim r As Long
        For r = 7 To hsLastRow
            Dim hsKP As String
            hsKP = Trim(CStr(hsWs.Cells(r, hsKPCol).Value))
            If hsKP <> "" And hsKP = kpNo Then
                foundRow = r
                foundByKP = True
                Exit For
            End If
        Next r
        
        ' KP-Noで見つからない場合は生産計画Noで検索（A列付近）
        If foundRow = 0 And seisanNo <> "" Then
            For r = 7 To hsLastRow
                Dim hsSeisan As String
                hsSeisan = Trim(CStr(hsWs.Cells(r, 1).Value))
                If hsSeisan = seisanNo Then
                    foundRow = r
                    Exit For
                End If
            Next r
        End If
        
        If foundRow = 0 Then
            ' 見つからない → 新規データ → 星取表末尾に追加
            Dim addRow As Long
            addRow = hsLastRow + 1
            hsWs.Cells(addRow, hsKPCol).Value = kpNo
            newCount = newCount + 1
            Call ログ書込("Step11_メンテ照合", "情報", _
                "行" & i & ": 新規データを星取表に追加（KP-No: " & kpNo & "）")
        Else
            ' 見つかった → 項目を比較
            Dim mismatch As String
            mismatch = ""
            
            ' 型式(G列)の比較 - 星取表のBH型式列
            Dim hsBHTypeCol As Long
            If InStr(model, "V8") > 0 Then
                hsBHTypeCol = 9  ' V8星取表のI列=BH型式
            Else
                hsBHTypeCol = 6  ' V9星取表のF列=BH型式
            End If
            Dim hsKatashiki As String
            hsKatashiki = Trim(CStr(hsWs.Cells(foundRow, hsBHTypeCol).Value))
            If katashiki <> "" And hsKatashiki <> "" And katashiki <> hsKatashiki Then
                mismatch = mismatch & "型式(" & katashiki & "≠" & hsKatashiki & ") "
            End If
            
            ' 出荷日の比較
            Dim hsDateCol As Long
            If InStr(model, "V8") > 0 Then
                hsDateCol = 11  ' V8星取表のK列=光真出荷日
            Else
                hsDateCol = 8   ' V9星取表のH列=光真出荷日
            End If
            Dim hsDate As Variant
            hsDate = hsWs.Cells(foundRow, hsDateCol).Value
            If IsDate(shukkaDate) And IsDate(hsDate) Then
                If CDate(shukkaDate) <> CDate(hsDate) Then
                    mismatch = mismatch & "出荷日(" & Format(CDate(shukkaDate), "M/D") & _
                        "≠" & Format(CDate(hsDate), "M/D") & ") "
                End If
            End If
            
            ' 順序指示発行日の比較（空欄は未確定なのでスキップ）
            If Not IsEmpty(junjoDate) And CStr(junjoDate) <> "" Then
                Dim hsJunjoCol As Long
                If InStr(model, "V8") > 0 Then
                    hsJunjoCol = 10  ' V8星取表のJ列
                Else
                    hsJunjoCol = 7   ' V9星取表のG列
                End If
                Dim hsJunjo As Variant
                hsJunjo = hsWs.Cells(foundRow, hsJunjoCol).Value
                If IsDate(junjoDate) And IsDate(hsJunjo) Then
                    If CDate(junjoDate) <> CDate(hsJunjo) Then
                        mismatch = mismatch & "発行日(" & Format(CDate(junjoDate), "M/D") & _
                            "≠" & Format(CDate(hsJunjo), "M/D") & ") "
                    End If
                End If
            End If
            
            ' 数量の比較（星取表では同じKP-Noの行数=台数）
            If IsNumeric(suryo) And Not IsEmpty(suryo) Then
                Dim hsRowCount As Long
                hsRowCount = 0
                Dim r2 As Long
                For r2 = 7 To hsLastRow
                    If Trim(CStr(hsWs.Cells(r2, hsKPCol).Value)) = kpNo Then
                        hsRowCount = hsRowCount + 1
                    End If
                Next r2
                If CLng(suryo) <> hsRowCount And hsRowCount > 0 Then
                    mismatch = mismatch & "数量(" & suryo & "≠" & hsRowCount & "行) "
                End If
            End If
            
            If mismatch = "" Then
                ' 全一致
                matchCount = matchCount + 1
            Else
                ' 一部一致 → 黄色ハイライト+ログ
                targetWs.Rows(i).Interior.Color = RGB(255, 255, 0)
                Call ログ書込("Step11_メンテ照合", "警告", _
                    "行" & i & ": 一部一致（要問い合わせ） KP-No:" & kpNo & " 不一致項目:" & mismatch)
                partialCount = partialCount + 1
            End If
        End If
NextRow:
    Next i
    
    ' ファイルを閉じる
    If Not v8Wb Is Nothing Then
        v8Wb.Save
        v8Wb.Close SaveChanges:=False
    End If
    If Not v9Wb Is Nothing And g_V9ProdSchedulePath <> g_V8ProdSchedulePath Then
        v9Wb.Save
        v9Wb.Close SaveChanges:=False
    End If
    
    ' 結果ログ
    Call ログ書込("Step11_メンテ照合", "完了", _
        "一致:" & matchCount & "件、一部一致:" & partialCount & "件、新規追加:" & newCount & "件")
    
    ' 一部一致があれば処理停止
    If partialCount > 0 Then
        MsgBox "【メンテ照合 - 要確認】" & vbCrLf & vbCrLf & _
               "一部一致が " & partialCount & " 件見つかりました。" & vbCrLf & _
               "黄色ハイライトされた行とログシートを確認し、" & vbCrLf & _
               "オムロン担当者に問い合わせてください。" & vbCrLf & vbCrLf & _
               "データ修正後、最初から再実行してください。", _
               vbExclamation, "メンテ照合 - 要問い合わせ"
        
        Application.ScreenUpdating = True
        Application.Calculation = xlCalculationAutomatic
        Application.EnableEvents = True
        End
    End If
End Sub

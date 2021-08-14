Sub 実行コマンド作成()
'
' 実行コマンド作成 Macro
' Keyboard Shortcut: Ctrl+q
'
    Dim ws As Worksheet
    Dim sqlFile As String
    Dim tableName As String
    Dim recordCount As Long
    Dim retu As Long
    Dim w_flg As Integer
    
    Const START_GYO As Integer = 5
    Const START_RETU As Integer = 6
    Const COL_NAME As Integer = 2  '項目名
    Const COL_TYPE_RETU As Integer = 4  '項目定義
    
    sqlFile = ActiveWorkbook.Path & "\" & Replace(ThisWorkbook.Name, ".xlsm", "") & "_" & Format(Now, "mmddhhnn") & ".sql"
    Open sqlFile For Output As #1
    
    'シートの個数分LOOP
    '(1シート目は決め打ちでスキップ。その為2〜)
    For sheet_i = 2 To Sheets.Count
        
        Set ws = ThisWorkbook.Sheets(sheet_i)
        
        ' =====================================================
        ' テーブルデータ作成用コマンド作成
        ' =====================================================
        If ws.Cells(1, 1).Value = "DML" Then
        
            Print #1, vbCrLf & "【DML】" & vbCrLf;
            
            'INSERTするレコード件数を判断
            For chk_i = 1 To 500
                If InStr(ws.Cells(4, chk_i).Value, "NOT NULL") > 0 Then
                    recordCount = chk_i - 1
                    Exit For
                End If
            Next
            
            'テーブル名取得
            tableName = ws.Cells(4, 2).Value
        
            Print #1, "DELETE FROM " & tableName & ";" & vbCrLf;
        
            For retu_i = START_RETU To recordCount
                
                gyo_i = START_GYO
                
                Do While ws.Cells(gyo_i, retu_i).Value <> "END"
                    '最終項目
                    If ws.Cells(gyo_i + 1, retu_i).Value = "END" Then
                        '設定値無し
                        If ws.Cells(gyo_i, retu_i).Value = "" Then
                                Print #1, "NULL);" & vbCrLf;
                        Else
                            If InStr(ws.Cells(gyo_i, COL_TYPE_RETU).Value, "CHAR") > 0 Or _
                               InStr(ws.Cells(gyo_i, COL_TYPE_RETU).Value, "TIMESTAMP") > 0 Then
                                Print #1, "'" & ws.Cells(gyo_i, retu_i).Value & "');" & vbCrLf;
                            Else
                                  Print #1, ws.Cells(gyo_i, retu_i).Value & ");" & vbCrLf;
                            End If
                        End If
    
                    '最終項目以外
                    Else
                        'テーブルの最初の項目
                        If gyo_i = START_GYO Then
                            '最初の項目が空だったら作成終了
                            If ws.Cells(gyo_i, retu_i).Value = "" Then
                                Exit For
                            Else
                                Print #1, "INSERT INTO " & tableName & " VALUES (";
                            End If
                        End If
                        
                        '設定値無し
                        If ws.Cells(gyo_i, retu_i).Value = "" Then
                            Print #1, "NULL,";
                        Else
                            If InStr(ws.Cells(gyo_i, COL_TYPE_RETU).Value, "CHAR") > 0 Or _
                               InStr(ws.Cells(gyo_i, COL_TYPE_RETU).Value, "TIMESTAMP") > 0 Then
                                Print #1, "'" & ws.Cells(gyo_i, retu_i).Value & "',";
                            Else
                                Print #1, ws.Cells(gyo_i, retu_i).Value & ",";
                            End If
                        End If
                    End If
                    gyo_i = gyo_i + 1
                    
                    If gyo_i > 200 Then
                        MsgBox "alert：[" & tableName & "]テーブル定義の最終行に「END」の記入忘れ？"
                        Print #1, "●エラー●" & vbCrLf;
                        Exit For
                    End If
                Loop
            Next
            Print #1, "COMMIT;" & vbCrLf;

        ' =====================================================
        ' APIリクエスト作成
        ' =====================================================
        Else
            
            Print #1, vbCrLf & "【APIリクエスト】" & vbCrLf;
            w_flg = 0
            
            'curlコマンドを作成件数を判断
            For chk_i = 1 To 500
                If InStr(ws.Cells(4, chk_i).Value, "NOT NULL") > 0 Then
                    recordCount = chk_i - 1
                    Exit For
                End If
            Next
            
            'システム機能ID取得
            ifName = LCase(ws.Cells(4, 2).Value)
        
            For retu_i = START_RETU To recordCount
                
                gyo_i = START_GYO
                
                Do While ws.Cells(gyo_i, retu_i).Value <> "END"
                    '最終項目
                    If ws.Cells(gyo_i + 1, retu_i).Value = "END" Then
                        '設定値あり
                        If ws.Cells(gyo_i, retu_i).Value <> "" Then
                                Print #1, "&" & ws.Cells(gyo_i, COL_NAME).Value & "=" & ws.Cells(gyo_i, retu_i).Value & "' http://localhost:28080/" & ifName & " | xmllint --format -";
                           Else
                                Print #1, "' http://localhost:XXXXX/" & ifName & " | xmllint --format -" & vbCrLf;
                        End If
    
                    '最終項目以外
                    Else
                        'テーブルの最初の項目
                        If gyo_i = START_GYO Then
                            '最初の項目が空だったら作成終了
                            If ws.Cells(gyo_i, retu_i).Value = "" Then
                                Exit For
                            Else
                                Print #1, "curl -X POST -d '";
                            End If
                        End If
                        
                        '設定値あり
                        If ws.Cells(gyo_i, retu_i).Value <> "" Then
                            If w_flg = 0 Then
                                Print #1, ws.Cells(gyo_i, COL_NAME).Value & "=" & ws.Cells(gyo_i, retu_i).Value;
                            Else
                                Print #1, "&" & ws.Cells(gyo_i, COL_NAME).Value & "=" & ws.Cells(gyo_i, retu_i).Value;
                            End If
                            w_flg = 1
                        End If
                    End If
                    gyo_i = gyo_i + 1
                    
                    If gyo_i > 200 Then
                        MsgBox "alert：[" & ifName & "]テーブル定義の最終行に「END」の記入忘れ？"
                        Print #1, "●エラー●" & vbCrLf;
                        Exit For
                    End If
                Loop
            Next
        End If
    Next
    
    Close #1
    
    MsgBox "info：完了"
    
    ' 処理完了後、出力先フォルダを開く
    'Shell "C:\Windows\Explorer.exe " & ActiveWorkbook.Path, vbNormalFocus

End Sub

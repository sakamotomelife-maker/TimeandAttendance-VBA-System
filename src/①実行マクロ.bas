Attribute VB_Name = "①実行マクロ"
Option Explicit

Sub 出勤登録ボタン()
       
    Application.ScreenUpdating = False
    On Error GoTo ErrorHandler
    
    Call 変数初期化
    Call 出勤登録準備
    Call 出勤登録エラー処理

    If err = 1 Then
            Exit Sub
   End If
       
    response = MsgBox(Split(staff_name, "　")(0) & "さん お疲れ様です！" & vbCrLf & vbCrLf & _
                Format(work_date, "YYYY/M/D") & "　" & Format(start_time, "h:mm") & "～" & Format(end_time, "h:mm") & vbCrLf & vbCrLf & _
                "上記内容で勤務しますか?", vbYesNo)
                  
    If response = vbYes Then
            
            
            Call 出勤交通転記   '---25/02/12
             
             wb.Save
            
             MsgBox "登録が完了しました。" & vbCrLf & vbCrLf & _
                        "不備がないか確認してください。"
                        
             Call 出勤登録リセット
             
             ThisWorkbook.Save
      
      Else
             MsgBox "処理を中止しました。"
      End If
      
      
ErrorHandler:
      
      Application.ScreenUpdating = True

End Sub


Sub 出勤取消ボタン()

      Application.ScreenUpdating = False
      Call 変数初期化
      
      If ws_toroku.Range("j6") = "" Then
                  MsgBox "出勤日が空白です。"
                  Exit Sub
      End If
      
      Call 出勤取消準備

      If err = 1 Then
                Exit Sub
       End If
        
    response = MsgBox(Split(staff_name, "　")(0) & "さん お疲れ様です！" & vbCrLf & vbCrLf & _
                Format(work_date, "YYYY/M/D") & " 勤務を取消しますか？", vbYesNo)
                  
                  
    If response = vbYes Then

            '出勤簿
            For i = 8 To 38
                  If Format(ws_syukkinbo.Cells(i, "A"), "D") = Format(work_date, "D") Then
                              ws_syukkinbo.Cells(i, "C").ClearContents
                              ws_syukkinbo.Cells(i, "e").ClearContents
                              ws_syukkinbo.Cells(i, "f").ClearContents
                              ws_syukkinbo.Cells(i, "h").ClearContents
                              ws_syukkinbo.Cells(i, "i").ClearContents
                              ws_syukkinbo.Cells(i, "k").ClearContents
                  End If
            Next i
                        
            '交通費明細書
            For i = 8 To 38
                   If Format(ws_kotsuhi.Cells(i, "A"), "D") = Format(work_date, "D") Then
                              ws_kotsuhi.Cells(i, "C").ClearContents
                              ws_kotsuhi.Cells(i, "E").ClearContents
                              ws_kotsuhi.Cells(i, "F").ClearContents
                              ws_kotsuhi.Cells(i, "G").ClearContents
                  End If
            Next i
            
             wb.Save
             'wb.Close                    '---wbを閉じずにその場で確認してもらう (25/1/9)

              MsgBox "取消が完了しました。" & vbCrLf & vbCrLf & _
                        "不備がないか確認してください。"
             
             Call 出勤取消リセット
      
      Else
             MsgBox "処理を中止しました。"
      End If

      Application.ScreenUpdating = True

End Sub

Sub 一括作成ボタン()

    Application.ScreenUpdating = False

    Call 変数初期化
    Call 支払票用シートリセット

    response = MsgBox("テンプレートから社員別の出勤簿エクセルを作成します。" & vbCrLf & _
                      "処理を実行しますか?", vbYesNo)

    If response = vbYes Then

              soshiki_no = ws_kanri.Range("C13")
              departmentName = ws_kanri.Range("F13")
      
              Dim new_name As String
      
              For i = 5 To last_info
      
                  'テンプレートエクセルを検索して開く
                  file_name = Dir(path_syukkin & "\*" & "組織番号　部署名　氏名" & "*.xlsx")
      
                  On Error GoTo ErrorHandler
                  If file_name <> "" Then
                  
                            Set wb = Workbooks.Open(path_syukkin & "\" & file_name)
                            Set ws_kotsuhi = wb.Sheets("交通費明細書")
                            Set ws_syukkinbo = wb.Sheets("出勤簿")
            
                            ' 社員情報の格納
                            staff_name = ws_kanri.Cells(i, "O").Value
                            phonotic = ws_kanri.Cells(i, "P").Value
                            joining_date = ws_kanri.Cells(i, "Q").Value
                            jikyu = ws_kanri.Cells(i, "R").Value
            
                            'テンプレ―トに入力
                            ws_syukkinbo.Range("C5") = soshiki_no
                            ws_syukkinbo.Range("d5") = departmentName
                            ws_syukkinbo.Range("e5") = staff_name
                            ws_syukkinbo.Range("f5") = phonotic
                            ws_syukkinbo.Range("G5") = joining_date
                            ws_syukkinbo.Range("k5") = jikyu
      
      
                            If InStrRev(file_name, "組織番号") > 0 And InStrRev(file_name, "氏名") > 0 Then
                                    new_name = Left(file_name, InStrRev(file_name, "組織番号") - 1) & _
                                                              soshiki_no & "　" & departmentName & "部　" & _
                                                              staff_name & Mid(file_name, InStrRev(file_name, "氏名") + Len("氏名"))
                                                          
                                    wb.SaveAs filename:=path_syukkin & "\" & new_name
                            End If
            
                            wb.Close
                      
                  Else
                      MsgBox "テンプレートファイルが見つかりませんでした。"
                      Exit Sub
                  End If
      
              Next i
      
              MsgBox "処理が完了しました。"
      
              Application.ScreenUpdating = True
              Exit Sub

ErrorHandler:
            MsgBox "各個別エクセルの作成元となるテンプレートファイルが見つかりませんでした。"
            On Error GoTo 0

    Else
            MsgBox "処理を中止しました。"
    End If

    Application.ScreenUpdating = True

End Sub



Sub 一括集計ボタン()

'1. #支払票用シートに集計
'2. 各シートの出勤簿と交通費を本エクセルにコピーする

    Application.ScreenUpdating = False

    Call 変数初期化
    Call 支払票用シートリセット
    
    
    response = MsgBox("一括集計を実行します。" & vbCrLf & _
                                 "既に実行済みの場合、前回作成分は削除されます。" & vbCrLf & vbCrLf & _
                                 "処理を実行しますか?", vbYesNo)
                                 
      If response = vbYes Then
      
                Call シート一括削除
      
                Dim j As Long
                Dim k As Long
                j = 2
            
            
                For i = 5 To last_info
                    
                        ' 社員情報の格納
                        staff_name = ws_kanri.Cells(i, "O").Value
                        phonotic = ws_kanri.Cells(i, "P").Value
                        joining_date = ws_kanri.Cells(i, "Q").Value
                        jikyu = ws_kanri.Cells(i, "R").Value
                        
                        ' 社員名でエクセルを検索して開く
                        file_name = Dir(path_syukkin & "\*" & staff_name & "*.xlsx")
                        
                        If file_name <> "" Then
                            
                                    ' 開いた社員別のエクセルに各変数セット
                                    Set wb = Workbooks.Open(path_syukkin & "\" & file_name)
                                    Set ws_kotsuhi = wb.Sheets("交通費明細書")
                                    Set ws_syukkinbo = wb.Sheets("出勤簿")
                                    
                                    Call 一括集計("出勤簿")
                                    Call 一括集計("交通費明細書")
                                    
                                    '#支払票用シートに値をコピー
                                    Dim sourceRange As Range
                                    Set sourceRange = wb.Sheets("支払票貼付用").Range("B2:M2")
                                    
                                    
                                    For k = 1 To sourceRange.Columns.Count
                                        ws_shiharai.Cells(j, "B").Offset(0, k - 1).Value = sourceRange.Cells(1, k).Value
                                    Next k
                                    
                                    j = j + 1
                                    
                                    wb.Save
                                    wb.Close
                        
                        Else
                                    MsgBox "ファイルが見つかりませんでした: " & staff_name
                        End If
                Next i
            
                  ws_kanri.Activate
                  MsgBox "処理が完了しました。"
      
      Else
                  MsgBox "処理を中止しました。"
      End If
      
      Application.ScreenUpdating = True
      
End Sub


Sub 支払票反映ボタン()

'1. 作成年月に基づいてシートを作成
'2. #支払票用シートの内容を支払票に反映

      Application.ScreenUpdating = False
      Call 変数初期化
      
      Dim shiharai_year As Long
      Dim shiharai_month As Long
      Dim sheetExists As Boolean
      Dim sheet_name As String
      
      shiharai_year = ws_kanri.Range("G23")
      shiharai_month = ws_kanri.Range("I23")
      
      response = MsgBox("'#支払票用' シートの内容で支払票を作成します。" & vbCrLf & _
                                    "※なお、" & shiharai_year & "年" & shiharai_month & "月分で作成します。" & vbCrLf & vbCrLf & _
                                    "処理を実行しますか？", vbYesNo)
      
          If response = vbYes Then
                  
                  '支払票を開く
                  file_name = Dir(path_shiharai & "\*支払*.xlsx")
                  On Error GoTo ErrorHandler
                  
                  Set wb = Workbooks.Open(path_shiharai & "\" & file_name)
      
                  sheet_name = "s" & Right(shiharai_year, 2) & "." & Format(shiharai_month, "00")
                  
                  
                  '既に対象分のシートが作られているのか確認
                  sheetExists = False
                  
                  For Each ws In wb.Sheets
                        If ws.Name = sheet_name Then
                              sheetExists = True
                              Exit For
                        End If
                  Next ws
             
                  If sheetExists = True Then
                  
                        response = MsgBox("支払票のエクセル内に対象年月のシートが存在しています。" & vbCrLf & _
                                                      "以前の入力内容を全て削除し、処理を続行してもよろしいですか？", vbYesNo)
                        
                        '前の入力を削除
                        If response = vbYes Then
                        
                              Set ws = wb.Sheets(sheet_name)
                             Call 支払票リセット
                        Else
                               MsgBox "処理を中止しました。"
                               Exit Sub
                              
                        End If
                        
                  Else
                         '新しいシート作成
                        wb.Sheets(wb.Sheets.Count).Copy after:=wb.Sheets(wb.Sheets.Count)
                        
                        Set ws = wb.Sheets(wb.Sheets.Count)
                        ws.Name = sheet_name
                        
                        Call 支払票リセット

                  End If
                        
                  '支払票用シートの内容をコピー
                  ws.Range("B9:j24").Value = ws_shiharai.Range("B2:j17").Value
                  ws.Range("L9:L24").Value = ws_shiharai.Range("L2:L17").Value
                  
                  ws.Range("j1") = Format(shiharai_month, "00")
                        
                  wb.Save

             MsgBox "処理が完了しました。" & vbCrLf & _
                        "※支払票は上書き保存済みです。"
             
             Call 出勤登録リセット
             Exit Sub
      
      Else
             MsgBox "処理を中止しました。"
             Exit Sub
      End If


ErrorHandler:
                  MsgBox "支払票が見つかりませんでした。" & vbCrLf & _
                                  "以下の指定ファイルパスに支払票が存在するか確認してください。" & vbCrLf & vbCrLf & _
                                  "●指定のファイルパス" & vbCrLf & _
                                  "<" & path_shiharai & "\" & file_name & ">"
            On Error GoTo 0
            
      Application.ScreenUpdating = True

End Sub


Sub 集計削除ボタン()

      Application.ScreenUpdating = False
      Call 変数初期化
      
      response = MsgBox("出勤簿と交通費明細書の各シートを削除します。" & vbCrLf & _
                                    "処理を実行しますか？", vbYesNo)
                                    
      If response = vbYes Then
      
            Call シート一括削除
            MsgBox "処理が完了しました。"
      Else
            MsgBox "処理を中止しました。"
      End If
      
      Application.ScreenUpdating = True
End Sub


Sub 格納ボタン()
      
    Application.ScreenUpdating = False
    Call 変数初期化
    
    Dim fso As Object
    Dim source As String
    Dim dest As String
    Dim file_name As String
    
    
     response = MsgBox("出勤簿・交通費明細書をアーカイブへ移動します。" & vbCrLf & _
                                  "※移動対象のエクセルは全て閉じてから実行してください。" & vbCrLf & vbCrLf & _
                                  "処理を実行しますか？", vbYesNo)
                              
      If response = vbYes Then
          
                      ' FileSystemObjectを作成
                      Set fso = CreateObject("Scripting.FileSystemObject")
                      
                      source = path_syukkin & "\"              ' 移動元
                      dest = path_syukkin & "\アーカイブ\"     ' 移動先
                      
                      
                      ' 移動先フォルダが存在しない場合は作成
                      If Not fso.FolderExists(dest) Then
                             fso.CreateFolder dest
                      End If
                      
                      ' 移動元フォルダ内のファイルを検索
                      file_name = Dir(source & "*.xlsx")
                      
                      Do While file_name <> ""
                              If InStr(file_name, "出勤簿・交通費明細") > 0 Then
                              
                                  ' 同名のファイルが存在する場合は削除
                                  If fso.FileExists(dest & file_name) Then
                                          fso.DeleteFile dest & file_name
                                  End If
                                  
                                  ' ファイルを移動
                                  fso.MoveFile source & file_name, dest & file_name
                              End If
                              
                              file_name = Dir
                      Loop
                  
                      MsgBox "処理が完了しました。"
                      Set fso = Nothing
                      
      Else
            MsgBox "処理を中止しました。"
      
      End If
      
    Application.ScreenUpdating = True

End Sub


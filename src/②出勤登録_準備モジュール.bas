Attribute VB_Name = "②出勤登録_準備モジュール"
Option Explicit

Sub 出勤登録エラー処理()

      If ws_toroku.Range("C8") = "" Or ws_toroku.Range("e8") = "" Then
                  MsgBox "時刻が入力されていないため処理を中止しました。"
                  
                  wb.Close
                  err = 1
                  Exit Sub
            End If
    
            If ws_toroku.Range("C8") > ws_toroku.Range("e8") Then
                  MsgBox "開始時刻が終了時刻よりも遅いため処理を中止しました。"
                  
                  wb.Close
                  err = 1

                  Exit Sub
            End If
            
            If ws_toroku.Range("AA6").Value <> 1 And ws_toroku.Range("e6") = "" Then
                  MsgBox "出勤日を入力する形式が指定されていますが、空欄のため処理を中止しました。"
            
                  wb.Close
                  err = 1
                  Exit Sub
            End If
            
End Sub


Sub 出勤登録準備()
    
    staff_name = ws_toroku.Range("C4").Value
    start_time = ws_toroku.Range("C8").Value
    end_time = ws_toroku.Range("E8").Value
    
    ' ファイルを検索して開く
    file_name = Dir(path_syukkin & "\*" & staff_name & "*.xlsx")
    
    If file_name <> "" And ws_toroku.Range("C4") <> "" Then
            Set wb = Workbooks.Open(path_syukkin & "\" & file_name)
    Else
            MsgBox "対象の出勤簿が見つかりませんでした。" & vbCrLf & _
                        "氏名に不備がないか確認の上、対象のエクセルが存在するか確認してください。"
            err = 1
            Exit Sub
    End If
    
    Set ws_kotsuhi = wb.Sheets("交通費明細書")
    Set ws_syukkinbo = wb.Sheets("出勤簿")
    
    
    ' 出勤日のラジオボタン
    If ws_toroku.Range("AA6").Value = 1 Then
            work_date = Date
    Else
            work_date = ws_toroku.Range("E6").Value
    End If


    work_day = Day(work_date)
    
    
    ' 社員情報の格納
    For i = 5 To last_info
            If staff_name = ws_kanri.Cells(i, "O").Value Then
                phonotic = ws_kanri.Cells(i, "P").Value
                joining_date = ws_kanri.Cells(i, "Q").Value
                jikyu = ws_kanri.Cells(i, "R").Value
                moyori = ws_kanri.Cells(i, "S").Value
                tocyaku = ws_kanri.Cells(i, "T").Value
                kotsuhi = ws_kanri.Cells(i, "U").Value
            End If
    Next i

End Sub


Sub 出勤取消準備()
    
    staff_name = ws_toroku.Range("j4").Value
        
    ' ファイルを検索して開く
    file_name = Dir(path_syukkin & "\*" & staff_name & "*.xlsx")
    
    If file_name <> "" And ws_toroku.Range("j4") <> "" Then
            Set wb = Workbooks.Open(path_syukkin & "\" & file_name)
    Else
            MsgBox "対象の出勤簿が見つかりませんでした。" & vbCrLf & _
                        "氏名に不備がないか確認の上、対象のエクセルが存在するか確認してください。"
            err = 1
            Exit Sub
    End If
    
    Set ws_kotsuhi = wb.Sheets("交通費明細書")
    Set ws_syukkinbo = wb.Sheets("出勤簿")

    work_date = ws_toroku.Range("J6")
    work_day = Day(work_date)

End Sub



Attribute VB_Name = "③一括集計モジュール"

Sub 一括集計(ByRef target As String)
            
            Dim ws_target As Worksheet
            Set ws_target = wb.Sheets(target)
            
            Dim name_option As String           'シート名が重複しないように付け足す記号用
            
            If target = "交通費明細書" Then
                        name_option = "●"
            Else
                        name_option = ""
            End If
      
             'シートコピー
              ws_target.Copy after:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)

              
              ' シート名が重複しないように確認
              Dim exists As Boolean
              exists = False
              Dim sh As Worksheet
              
              For Each sh In ThisWorkbook.Sheets
                  
                  If sh.Name = staff_name & name_option Then
                      exists = True
                      Exit For
                  End If
              Next sh
              
             'シート名を変更
              If exists Then
                  MsgBox "シート名が重複しています: " & staff_name
              Else
                  ThisWorkbook.Sheets(target).Name = staff_name & name_option
              End If
            
End Sub

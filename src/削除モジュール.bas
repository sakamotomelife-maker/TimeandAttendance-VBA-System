Attribute VB_Name = "削除モジュール"

Sub 出勤登録リセット()

      ws_toroku.Range("C4:E4").ClearContents
      ws_toroku.Range("C8").ClearContents
      ws_toroku.Range("e8").ClearContents
      ws_toroku.Range("AA6") = 1

End Sub


Sub 出勤取消リセット()

      ws_toroku.Range("j4:L4").ClearContents

End Sub


Sub 支払票用シートリセット()

      ws_shiharai.Range("b2:m11").ClearContents
      
End Sub


Sub 支払票リセット()

      ws.Range("B9:J24").ClearContents
      ws.Range("L9:L24").ClearContents
      
End Sub


'元々のシート以外まとめて削除
Sub シート一括削除()

    ' 各シートをチェック
    For Each ws In ThisWorkbook.Sheets
    
        ' シート名が「一覧」、「テンプレート」、「格納」以外の場合、シートを削除
        If ws.Name <> "#出勤登録" And ws.Name <> "#管　　理" And ws.Name <> "#支払票用" Then
        
                  Application.DisplayAlerts = False
                  
                  ws.Delete
                  
                  Application.DisplayAlerts = True

        End If
        
    Next ws
    
End Sub

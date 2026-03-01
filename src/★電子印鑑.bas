Attribute VB_Name = "★電子印鑑"


Sub 押印実行ボタン()

    Application.ScreenUpdating = False
    Call 変数初期化
    Dim j As Long

    '本シートに電子印を押す場合
    If ws_kanri.Range("aa25") = 1 Then

                 response = MsgBox("本エクセルに集約した各シートに押印します。" & vbCrLf & _
                                                "処理を実行しますか？", vbYesNo)

                  If response = vbYes Then
                        Call 本シート押印
                  Else
                        MsgBox "処理を中止しました。"
                        Exit Sub
                  End If

      '個別のエクセルに電子印を押す場合
      Else
           response = MsgBox("社員別の各エクセルに押印します。" & vbCrLf & _
                                                "処理を実行してもよろしいですか？", vbYesNo)

                  If response = vbYes Then

                          For i = 5 To last_info

                                  ' 社員情報の格納
                                   staff_name = ws_kanri.Cells(i, "O").Value

                                  ' 社員名でエクセルを検索して開く
                                  file_name = Dir(path_syukkin & "\*" & staff_name & "*.xlsx")

                                  If file_name <> "" Then

                                      ' 開いた社員別のエクセルに各変数セット
                                        Set wb = Workbooks.Open(path_syukkin & "\" & file_name)
                                        Set ws_kotsuhi = wb.Sheets("交通費明細書")
                                        Set ws_syukkinbo = wb.Sheets("出勤簿")

                                  Else
                                        MsgBox "ファイルが見つかりませんでした: " & staff_name
                                        Exit Sub
                                  End If

                                 Call 個別シート押印(ws_kotsuhi, ws_syukkinbo)
                                 Call 画像位置調整(ws_kotsuhi, ws_syukkinbo, 20, 0)
                                 
                                 wb.Save
                                 wb.Close
                                 
                        Next i

                  Else
                              MsgBox "処理を中止しました。"
                  End If
      End If

       MsgBox "処理が完了しました。"
       ws_kanri.Activate

      Application.ScreenUpdating = True

End Sub


Sub 押印削除ボタン()

      Application.ScreenUpdating = False
      Call 変数初期化

      '本シートの電子印を消す場合
       If ws_kanri.Range("aa25") = 1 Then

                 response = MsgBox("本エクセルの押印を削除します。" & vbCrLf & _
                                                "処理を実行しますか？", vbYesNo)

                  If response = vbYes Then

                            For Each ws In ThisWorkbook.Sheets
                            
                                    If InStr(ws.Name, "#") = 0 Then
                                              Call delete_shape(ws)
                                    End If
                            Next ws

                  Else
                        MsgBox "処理を中止しました。"
                        Exit Sub
                  End If

      '個別シートの電子印を消す場合の処理
      ElseIf ws_kanri.Range("aa25") = 2 Then
                  response = MsgBox("社員別の各エクセルの押印を削除します。" & vbCrLf & _
                                                "処理を実行してもよろしいですか？", vbYesNo)

                  If response = vbYes Then

                          For i = 5 To last_info

                                  ' 社員情報の格納
                                   staff_name = ws_kanri.Cells(i, "O").Value

                                  ' 社員名でエクセルを検索して開く
                                  file_name = Dir(path_syukkin & "\*" & staff_name & "*.xlsx")

                                  If file_name <> "" Then

                                      ' 開いた社員別のエクセルに各変数セット
                                        Set wb = Workbooks.Open(path_syukkin & "\" & file_name)
                                        Set ws_kotsuhi = wb.Sheets("交通費明細書")
                                        Set ws_syukkinbo = wb.Sheets("出勤簿")

                                  Else
                                        MsgBox "ファイルが見つかりませんでした: " & staff_name
                                        Exit Sub
                                  End If

                                 Call delete_shape(ws_kotsuhi)
                                 Call delete_shape(ws_syukkinbo)
                                 
                                 wb.Save
                                 wb.Close
                                                                  
                        Next i

                  Else
                              MsgBox "処理を中止しました。"
                              Exit Sub
                  End If
      End If


      MsgBox "処理が完了しました。"
      Application.ScreenUpdating = True

End Sub


Sub 個別エクセル押印調整ボタン()

      Application.ScreenUpdating = False

      Call 変数初期化
      Dim move_setting As String
      
              For i = 5 To last_info

                      ' 社員情報の格納
                       staff_name = ws_kanri.Cells(i, "O").Value

                      ' 社員名でエクセルを検索して開く
                      file_name = Dir(path_syukkin & "\*" & staff_name & "*.xlsx")

                      If file_name <> "" Then

                          ' 開いた社員別のエクセルに各変数セット
                            Set wb = Workbooks.Open(path_syukkin & "\" & file_name)
                            Set ws_kotsuhi = wb.Sheets("交通費明細書")
                            Set ws_syukkinbo = wb.Sheets("出勤簿")

                      Else
                            MsgBox "ファイルが見つかりませんでした。: " & staff_name
                            Exit Sub
                      End If

                     Call 画像位置調整(ws_kotsuhi, ws_syukkinbo, ws_kanri.Range("aa28"), 20, 0)
                     
                     wb.Save
                     wb.Close
                     
            Next i

       MsgBox "処理が完了しました。"
       ws_kanri.Activate

      Application.ScreenUpdating = True

End Sub

'本シート押印は本モジュールを使用する
Sub 押印(ByRef ws As Worksheet, ByRef target As Range)

    Dim stamp As Shape

    ' 電子印鑑の形状を取得
    Set stamp = ws_kanri.Shapes("Stamp")

    ' 対象のセルに電子印鑑を貼り付け
    stamp.Copy
    ws.Paste Destination:=target

End Sub

 
 '---個別エクセルでシートが保護されているため実行できないというエラー発生したため改良ver(24/1/9)
Sub 押印v2(ByRef ws As Worksheet, ByRef target As Range)

    Dim stamp As Range

    ' 電子印鑑の形状を取得
    Set stamp = ws_kanri.Range("j13")

    ' 対象のセルに電子印鑑を貼り付け
    If Not target.Locked Then
            stamp.Copy
            ws.Paste Destination:=target
      End If


End Sub


'図形の削除
Sub delete_shape(ByRef ws_target As Worksheet)

    Dim shp As Shape

    For Each shp In ws_target.Shapes
            shp.Delete
    Next shp

End Sub


Sub 本シート押印()

        For Each ws In ThisWorkbook.Sheets

                  '出勤簿
                  If ws.Range("A1").Value = "サンプル出勤簿" Then

                            For i = 8 To 38
                                        If ws.Cells(i, "L").Value > 0 Then
                                            Call 押印(ws, ws.Cells(i, "M"))
                                        End If
                            Next i

                  '交通費
                  ElseIf ws.Range("A1").Value = "サンプル交通費明細書" Then

                            For i = 8 To 38
                                    If ws.Cells(i, "R").Value > 0 Then
                                        Call 押印(ws, ws.Cells(i, "S"))
                                    End If
                            Next i
                  End If
      Next ws

End Sub


Sub 個別シート押印(ByRef ws_kotsuhi As Worksheet, ByRef ws_syukkinbo As Worksheet)

                              For j = 8 To 38
                                                                
                                  If ws_syukkinbo.Cells(j, "L").Value > 0 Then
                                      Call 押印v2(ws_syukkinbo, ws_syukkinbo.Cells(j, "M"))
                                  End If
                              Next j

                              For j = 8 To 38

                                  If ws_kotsuhi.Cells(j, "R").Value > 0 Then
                                      Call 押印v2(ws_kotsuhi, ws_kotsuhi.Cells(j, "S"))
                                  End If
                              Next j

End Sub

'---個別エクセルに押印すると位置がずれるため使用 (24/01/11)
Sub 画像位置調整(ByRef ws_kotsuhi As Worksheet, ByRef ws_syukkinbo As Worksheet, ByRef move_top As Long, ByRef move_left As Long)

      Dim img As Shape
      
      '出勤簿
      For Each img In ws_syukkinbo.Shapes
            img.Top = img.Top + move_top
            img.Left = img.Left + move_left
      Next img
      
      '交通費
      For Each img In ws_kotsuhi.Shapes
            img.Top = img.Top + move_top
            img.Left = img.Left + move_left
      Next img
            
End Sub

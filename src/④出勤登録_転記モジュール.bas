Attribute VB_Name = "④出勤登録_転記モジュール"
Option Explicit

'交通費転記と出勤転記のfor文をまとめる(24/02/12)
Sub 出勤交通転記()

Dim date_line As Long   '入力する行番号を格納

'出勤転記
For i = 8 To 38
                       If Day(ws_syukkinbo.Cells(i, "A").Value) = work_day Then
                                    
                                date_line = i
                        
                        
                                  '勤務開始---------------------------------------------------------------
                                  
                                  '午前勤務
                                  If start_time < TimeValue("12:00:00") Then
                                            ws_syukkinbo.Cells(i, "C").Value = start_time
                                  
                                  '午後勤務
                                  ElseIf start_time = TimeValue("14:00:00") Then
                                            ws_syukkinbo.Cells(i, "F").Value = start_time
                              
                                  '夜勤務
                                  ElseIf start_time > TimeValue("18:00:00") Then
                                            ws_syukkinbo.Cells(i, "I").Value = start_time
                                  End If
                              
                              
                                  '午前終了---------------------------------------------------------------
                                  
                                  '12:00まで
                                  If end_time = TimeValue("12:00:00") Then
                                            ws_syukkinbo.Cells(i, "E").Value = end_time
                                            
                                  '午前に出勤していて13:00迄または午後も働く場合
                                  ElseIf start_time < TimeValue("12:00:00") And end_time > TimeValue("12:00:00") Then
                                             ws_syukkinbo.Cells(i, "E").Value = TimeValue("13:00")
                                  End If
                              
                              
                                  '午後終了---------------------------------------------------------------
                                  
                                  '16:00まで
                                  If end_time = TimeValue("16:00:00") Then
                                            ws_syukkinbo.Cells(i, "H").Value = end_time
                                            
                                  '午前に出勤していて17:00迄または午後も働く場合
                                  ElseIf end_time >= TimeValue("17:00:00") And start_time < TimeValue("17:00:00") Then
                                             ws_syukkinbo.Cells(i, "H").Value = TimeValue("17:00")
                                  End If
                                  
                                  
                                  '夜終了---------------------------------------------------------------
                                  
                                  '20:30まで
                                  If end_time = TimeValue("20:30") Then
                                            ws_syukkinbo.Cells(i, "K").Value = end_time
                                  End If
                                  
                                  
                                  '休憩後の勤務開始-----------------------------------------------------
                                    
                                    '昼休憩
                                    If ws_syukkinbo.Cells(i, "H").Value <> "" And ws_syukkinbo.Cells(i, "C").Value <> "" Then
                                            ws_syukkinbo.Cells(i, "F").Value = TimeValue("14:00:00")
                                    End If
                                    
                                    '夜休憩
                                    If ws_syukkinbo.Cells(i, "K").Value <> "" And ws_syukkinbo.Cells(i, "F").Value <> "" Then
                                            ws_syukkinbo.Cells(i, "I").Value = TimeValue("18:30:00")
                                    End If
                          Exit For
                          
                          End If
                Next i

'交通費転記
                              
                              i = date_line
                              
                              kotsu_raw = i
                              ws_kotsuhi.Cells(i, "C").Value = moyori
                              ws_kotsuhi.Cells(i, "E").Value = tocyaku
                              
                              If ws_kotsuhi.Cells(i, "C").Value <> "" Then
                                       ws_kotsuhi.Cells(i, "F").Value = "往復"
                              End If
                              
                              ws_kotsuhi.Cells(i, "G").Value = kotsuhi
                              
                


End Sub



Sub 交通費転記()

 For i = 8 To 38
                    If Day(ws_kotsuhi.Cells(i, "A").Value) = work_day Then
                    
                              kotsu_raw = i
                              ws_kotsuhi.Cells(i, "C").Value = moyori
                              ws_kotsuhi.Cells(i, "E").Value = tocyaku
                              
                              If ws_kotsuhi.Cells(i, "C").Value <> "" Then
                                       ws_kotsuhi.Cells(i, "F").Value = "往復"
                              End If
                              
                              ws_kotsuhi.Cells(i, "G").Value = kotsuhi
                              
                              Exit Sub
                    End If
                Next i
End Sub


Sub 出勤転記()

  For i = 8 To 38
                       If Day(ws_syukkinbo.Cells(i, "A").Value) = work_day Then
                                    
                        
                                  '勤務開始---------------------------------------------------------------
                                  
                                  '午前勤務
                                  If start_time < TimeValue("12:00:00") Then
                                            ws_syukkinbo.Cells(i, "C").Value = start_time
                                  
                                  '午後勤務
                                  ElseIf start_time = TimeValue("14:00:00") Then
                                            ws_syukkinbo.Cells(i, "F").Value = start_time
                              
                                  '夜勤務
                                  ElseIf start_time > TimeValue("18:00:00") Then
                                            ws_syukkinbo.Cells(i, "I").Value = start_time
                                  End If
                              
                              
                                  '午前終了---------------------------------------------------------------
                                  
                                  '12:00まで
                                  If end_time = TimeValue("12:00:00") Then
                                            ws_syukkinbo.Cells(i, "E").Value = end_time
                                            
                                  '午前に出勤していて13:00迄または午後も働く場合
                                  ElseIf start_time < TimeValue("12:00:00") And end_time > TimeValue("12:00:00") Then
                                             ws_syukkinbo.Cells(i, "E").Value = TimeValue("13:00")
                                  End If
                              
                              
                                  '午後終了---------------------------------------------------------------
                                  
                                  '16:00まで
                                  If end_time = TimeValue("16:00:00") Then
                                            ws_syukkinbo.Cells(i, "H").Value = end_time
                                            
                                  '午前に出勤していて17:00迄または午後も働く場合
                                  ElseIf end_time >= TimeValue("17:00:00") And start_time < TimeValue("17:00:00") Then
                                             ws_syukkinbo.Cells(i, "H").Value = TimeValue("17:00")
                                  End If
                                  
                                  
                                  '夜終了---------------------------------------------------------------
                                  
                                  '20:30まで
                                  If end_time = TimeValue("20:30") Then
                                            ws_syukkinbo.Cells(i, "K").Value = end_time
                                  End If
                                  
                                  
                                  '休憩後の勤務開始-----------------------------------------------------
                                    
                                    '昼休憩
                                    If ws_syukkinbo.Cells(i, "H").Value <> "" And ws_syukkinbo.Cells(i, "C").Value <> "" Then
                                            ws_syukkinbo.Cells(i, "F").Value = TimeValue("14:00:00")
                                    End If
                                    
                                    '夜休憩
                                    If ws_syukkinbo.Cells(i, "K").Value <> "" And ws_syukkinbo.Cells(i, "F").Value <> "" Then
                                            ws_syukkinbo.Cells(i, "I").Value = TimeValue("18:30:00")
                                    End If
                          Exit Sub
                          
                          End If
                Next i
End Sub

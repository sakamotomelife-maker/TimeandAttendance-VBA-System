Attribute VB_Name = "★印刷マクロ"

Sub 一括PDFボタン()

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    Call 変数初期化
    Dim sheetArray() As String

    response = MsgBox("1つのPDFにまとめて出力します。" & vbCrLf & _
                      "処理を実行しますか？", vbYesNo)

    If response = vbYes Then

        ' シート名を配列に格納
        i = 0

        For Each ws In ThisWorkbook.Sheets

            If ws.Name <> "#出勤登録" And ws.Name <> "#管　　理" And ws.Name <> "#支払票用" Then

                ReDim Preserve sheetArray(i)
                sheetArray(i) = ws.Name
                i = i + 1

            End If
        Next ws

        'err
        If i = 0 Then
            MsgBox "印刷対象のシートが存在しません。"
            Exit Sub
        End If

        ' シートを選択
        ThisWorkbook.Sheets(sheetArray).Select

        ' 印刷設定を適用
        For Each ws In ThisWorkbook.Sheets(sheetArray)
                  With ws.PageSetup
                        .CenterHorizontally = True
                        .CenterVertically = True
                        .TopMargin = Application.InchesToPoints(0.5)
                        .BottomMargin = Application.InchesToPoints(0.5)
                        .LeftMargin = Application.InchesToPoints(0.5)
                        .RightMargin = Application.InchesToPoints(0.5)
                        .Zoom = False
                        .FitToPagesWide = 1
                        .FitToPagesTall = 1
                  End With
        Next ws

'-----------------------------------------------------------------------------------------------------------------------

        ' PDFとして保存
        Dim pdfFileName As String
        Dim path_pdf As String
        
        pdfFileName = "XXXXXX　XX月度出勤簿　" & soshiki_no & "　" & departmentName & "部.pdf"
        path_pdf = path_shiharai & "\"
      
        ' ファイルが存在するか確認
        If Dir(path_pdf & pdfFileName) <> "" Then
                  Dim cnt As Long
                  cnt = 1
                  Do While Dir(path_pdf & pdfFileName) <> ""
                      pdfFileName = "XXXXXX　XX月度出勤簿　" & soshiki_no & "　" & departmentName & "部_" & cnt & ".pdf"
                      cnt = cnt + 1
                  Loop
        End If
            
        ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, _
                                        filename:=path_pdf & pdfFileName, _
                                        Quality:=xlQualityStandard, _
                                        IncludeDocProperties:=True, _
                                        IgnorePrintAreas:=False, _
                                        OpenAfterPublish:=True

        ' シート選択を解除
        ThisWorkbook.Sheets("#管　　理").Select

        MsgBox "処理を完了しました。" & vbCrLf & _
               "※支払票と同じフォルダパスに保存しています。"

    Else
        MsgBox "処理を中止しました。"
    End If

    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True

End Sub



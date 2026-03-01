Attribute VB_Name = "変数定義"
Option Explicit

Global wb As Workbook
Global ws As Worksheet
Global ws_toroku As Worksheet
Global ws_kotsuhi As Worksheet
Global ws_kanri As Worksheet
Global ws_syukkinbo As Worksheet
Global ws_shiharai As Worksheet

Global staff_name As String
Global work_date As Date
Global work_day As String
Global start_time As Date
Global end_time As Date

Global i As Long
Global kotsu_raw As Long

Global phonotic As String
Global joining_date
Global jikyu As Long
Global moyori As String
Global tocyaku As String
Global kotsuhi As String

Global last_info As Long
Global last_shiharai As Long

Global path_syukkin As String
Global path_shiharai As String
Global file_name As String

Global soshiki_no As Long
Global departmentName As String

'エラー処理用
Global response As VbMsgBoxResult
Global err As Long


Sub 変数初期化()

    err = 0
    
    Set ws_toroku = ThisWorkbook.Sheets("#出勤登録")
    Set ws_kanri = ThisWorkbook.Sheets("#管　　理")
    Set ws_shiharai = ThisWorkbook.Sheets("#支払票用")
    
    path_syukkin = ws_kanri.Range("E5").Value
    path_shiharai = ws_kanri.Range("E7").Value
    
    soshiki_no = ws_kanri.Range("C13").Value
    departmentName = ws_kanri.Range("F13").Value
    
    
    last_info = ws_kanri.Cells(ws_kanri.Rows.Count, "O").End(xlUp).Row
    last_shiharai = ws_shiharai.Cells(ws_shiharai.Rows.Count, "B").End(xlUp).Row
        
End Sub




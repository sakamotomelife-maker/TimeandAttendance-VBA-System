VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "個別エクセル位置補正 - 電子印鑑"
   ClientHeight    =   2856
   ClientLeft      =   -132
   ClientTop       =   -780
   ClientWidth     =   7644
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Image2_BeforeDragOver(ByVal Cancel As MSForms.ReturnBoolean, ByVal Data As MSForms.DataObject, ByVal X As Single, ByVal Y As Single, ByVal DragState As MSForms.fmDragState, ByVal Effect As MSForms.ReturnEffect, ByVal Shift As Integer)
      
            
End Sub


'中止ボタン
Private Sub CommandButton2_Click()

      Unload UserForm1

End Sub

'適用ボタン
Private Sub CommandButton3_Click()

      Call 個別エクセル押印調整ボタン
      
      ActiveSheet.Range("AA28").Value = TextBox47
      
      Unload UserForm1

End Sub


Private Sub Image1_BeforeDragOver(ByVal Cancel As MSForms.ReturnBoolean, ByVal Data As MSForms.DataObject, ByVal X As Single, ByVal Y As Single, ByVal DragState As MSForms.fmDragState, ByVal Effect As MSForms.ReturnEffect, ByVal Shift As Integer)

End Sub

Private Sub Label9_Click()

End Sub

Private Sub SpinButton1_Change()
      
      TextBox47 = ActiveSheet.Range("AA28").Value + SpinButton1.Value
      
End Sub


Private Sub TextBox47_Change()

End Sub

Private Sub UserForm_Initialize()
      
      TextBox47.MaxLength = 3
      TextBox47 = ActiveSheet.Range("AA28").Value
      
      SpinButton1.Max = 100 - ActiveSheet.Range("AA28").Value
      SpinButton1.Min = -100 - ActiveSheet.Range("AA28").Value
      
End Sub

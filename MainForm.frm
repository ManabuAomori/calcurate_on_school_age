VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} MainForm 
   Caption         =   "日付選択"
   ClientHeight    =   2520
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3945
   OleObjectBlob   =   "MainForm.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "MainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmnd_btn_Click()
If TextBox1.Value > TextBox2.Value Then
    MsgBox "誕生日以降の日付を判定月日に指定してください"
    Exit Sub
End If
TextBox3.Text = Gakunerei(CDate(TextBox1.Value), CDate(TextBox2.Value))
End Sub

Private Sub CommandButton1_Click()
  Call ShowCalender(1)
End Sub
 
Private Sub CommandButton2_Click()
  Call ShowCalender(2)
End Sub
 
Private Sub CommandButton3_Click()
  Call ShowCalender(3)
End Sub

Private Sub ShowCalender(i As Integer)
  clndr_flg = False 'フラグリセット
  If IsDate(Me("TextBox" & i).Value) = False Then '日付が入ってなければ
    clndr_date = Date '今日の日付を格納
  Else
    clndr_date = Me("TextBox" & i).Value 'テキストボックスの日付を格納
  End If
  CalenderForm.Show 'カレンダーを開く
  If clndr_flg = True Then Me("TextBox" & i).Value = Format(clndr_date, "yyyy/mm/dd") 'クリックされた日付を上書き
End Sub

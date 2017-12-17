Attribute VB_Name = "Module1"
Option Explicit '変数の宣言を強制する

'---カレンダー用変数
Public clndr_date As Date 'テキストボックスの値を格納する変数
Public clndr_flg As Boolean 'カレンダーがクリックされたか判定するフラグ

Sub start()
  MainForm.Show 'メインフォームを開く
End Sub
'######
'# Calcurating school years
'######
Public Function Gakunerei(birthDay As Date, orderDay As Date) As Variant


Dim tempDay1, tempMonth, r_value As Variant

    If IsDate(birthDay) Then
         If Format(birthDay, "yyyy/mm/dd") >= Format(Year(birthDay) & "/01/01", "yyyy/mm/dd") And Format(birthDay, "yyyy/mm/dd") <= Format(Year(birthDay) & "/04/01", "yyyy/mm/dd") Then
            tempDay1 = DateAdd("yyyy", -1, birthDay)
            tempMonth = DateDiff("m", tempDay1, orderDay)
        Else
            If Year(birthDay) < Year(orderDay) Then
            tempDay1 = DateAdd("yyyy", -1, birthDay)
            tempMonth = DateDiff("m", tempDay1, orderDay)
            Else
            tempMonth = DateDiff("m", birthDay, orderDay)
            End If
        End If
    Else
        Exit Function
    End If
    
    r_value = ((Int(tempMonth / 12) * 100) + (tempMonth Mod 12)) / 100
    
    
   If Not r_value = "" Then
        Gakunerei = r_value
    Else
        Gakunerei = -1
    End If
    
End Function

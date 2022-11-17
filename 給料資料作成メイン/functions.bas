Attribute VB_Name = "functions"
'// シートで使用する関数を定義するモジュール
Option Explicit

'// シートで使用する関数を関数ダイアログボックスに登録
Public Sub registerFunctions()

    Application.MacroOptions _
        Macro:="PAYMENT", Description:="項目(例:社会保険料)と対象部署から金額を求める関数です。", _
        Category:="給料資料作成", ArgumentDescriptions:=Array("項目(例:社会保険料)", "対象部署コードまたは部署名", "シート名が勤怠支給控除一覧表の場合は省略可能です｡ ")

End Sub

'/**
' * 控除金額を「勤怠支給控除一覧表」から計算
' * @params targetType 控除項目(健康保険料など)
' * @params targetCode 控除対象部署のコードもしくは名前

Public Function PAYMENT(ByVal targetType As Variant, targetCode As Variant, Optional sheetName As String = "勤怠支給控除一覧表") As Long
Attribute PAYMENT.VB_Description = "項目(例:社会保険料)と対象部署から金額を求める関数です。"
Attribute PAYMENT.VB_ProcData.VB_Invoke_Func = " \n20"

    Application.Volatile

   On Error Resume Next
    
    '// 項目の行番号
    Dim targetRow As Long
    
    If sheetName = "勤怠支給控除一覧表" Then
        targetRow = WorksheetFunction.Match("*" & targetType & "*", Sheets(sheetName).Columns(1), 0)
    ElseIf sheetName = "控除" Then
        targetRow = WorksheetFunction.Match("*" & targetCode & "*", Sheets(sheetName).Columns(2), 0)
    End If
    
    '// 対象部署の列番号
    Dim targetColumn As Long
    
    If sheetName = "勤怠支給控除一覧表" Then
        targetColumn = WorksheetFunction.Match("*" & targetCode & "*", Sheets(sheetName).Rows(5), 0)
    ElseIf sheetName = "控除" Then
        targetColumn = WorksheetFunction.Match("*" & targetType & "*", Sheets(sheetName).Rows(5), 0)
    End If
    
    On Error GoTo 0
    
    '// 控除項目・控除対象部署が見つからない、もしくはそのセルの値が数字出ない場合は抜ける
    If targetRow = 0 Or targetColumn = 0 Then
        PAYMENT = 0
        Exit Function
    ElseIf IsNumeric(Sheets(sheetName).Cells(targetRow, targetColumn).value) = False Then
        PAYMENT = 0
        Exit Function
    End If
        
    PAYMENT = Sheets(sheetName).Cells(targetRow, targetColumn).value

End Function
'// 奉行から出力されたデータの日付から支払日を計算
Public Function PAYDAY(ByVal dateOfBugyo As Variant) As Date
    
    Application.Volatile
    
    '1 奉行の日付を"年" でわけ、年と月を求める
    dateOfBugyo = Split(dateOfBugyo, "年")
    
    '// 和暦を西暦に変換
    Dim yearOfBugyo As Long: yearOfBugyo = Format(dateOfBugyo(0) & "年1月1日", "yyyy")
    '// 月を取得
    Dim monthOfbugyo As Long: monthOfbugyo = Val(Split(dateOfBugyo(1), "月")(0))
    dateOfBugyo = DateSerial(yearOfBugyo, monthOfbugyo, 20)
    
    '2 土日祝日と重なったら日付を1日前にして、平日になるまで繰り替えす
    Dim result As Boolean, i As Long
    i = 1
    Do Until result = True
        If Weekday(dateOfBugyo) = 1 Or Weekday(dateOfBugyo) = 7 Then
            dateOfBugyo = DateSerial(yearOfBugyo, monthOfbugyo, 20 - i)
            result = False
            i = i + 1
        ElseIf Application.WorksheetFunction.CountIf(Sheets("設定").Range("A:A"), dateOfBugyo) >= 1 Then
            dateOfBugyo = DateSerial(yearOfBugyo, monthOfbugyo, 20 - i)
            result = False
            i = i + 1
        Else
            result = True
        End If
    Loop
    PAYDAY = dateOfBugyo

End Function

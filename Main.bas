Attribute VB_Name = "Main"
'// 給料資料作成のメインモジュール
Option Explicit

'// 表加工のためのフォーム起動
Public Sub openFormToProcessChart()

    Sheets("mode").Cells(1, 1).Value = "PROCESS_CHART"
    frmCompany.Show
    
End Sub

'// メインルーチン(表加工)
Public Sub processChart(ByVal company As String)

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    '// 途中でエラーになった時に貼り付けなおしにならないように表の最初の状態を他のシートに保持する
    Sheets("勤怠支給控除一覧表").Cells.Copy Sheets("tmp").Cells(1, 1)
    
    Dim companyRow As Long: companyRow = WorksheetFunction.Match(company, Sheets("設定").Columns(1), 0)
    Dim headerRow As Long: headerRow = Sheets("勤怠支給控除一覧表").Cells(1, 3).End(xlDown).Row

    Dim cc As New ChartController
    Dim vc As New ValueController

    '// 不要な部署を削除
    If deleteUnnecessaryColumns(companyRow, headerRow, cc) = False Then
        Call resetChart
        GoTo Kill
    End If
    
    '// 「現場」にも「事務」にも含まない列移動
    If moveIndependentColumns(companyRow, headerRow, cc) = False Then
        Call resetChart
        GoTo Kill
    End If
    
    '// 事務列作成
    If createOfficeColumn(companyRow, headerRow, cc, vc) = False Then
        Call resetChart
        GoTo Kill
    End If
    
    '// 現場列作成
    If createFieldWorkColumn(companyRow, headerRow, cc) = False Then
        Call resetChart
        GoTo Kill
    End If
    '// 支給金額行作成
    Call createBasicSalaryRow(headerRow)
    
    '// 勤怠支給控除一覧表の見た目調整
    With Sheets("勤怠支給控除一覧表")
        Dim mainLastColumn As Long
        
        If Sheets("設定").Cells(companyRow, 3).Value = "" And Sheets("設定").Cells(companyRow, 4).Value = "" Then
            mainLastColumn = 0
        ElseIf Sheets("設定").Cells(companyRow, 3).Value <> "" And Sheets("設定").Cells(companyRow, 4).Value = "" Then
            mainLastColumn = WorksheetFunction.Match("事務", .Rows(headerRow), 0)
        Else
            mainLastColumn = cc.searchDepartmentColumn(Int(Split(Sheets("設定").Cells(2, 4).Value, "-")(0)), 5)
        End If
        
        If mainLastColumn <> 0 Then
            .Range(Cells(headerRow, 1), .Cells(.Cells(Rows.Count, 1).End(xlUp).Row, mainLastColumn)).BorderAround Weight:=xlMedium
        End If
        
        .Range(.Rows(5), .Rows(6)).HorizontalAlignment = xlCenter
        
    End With
    
    '// 控除の表修正
    Call processDeductionChart(vc)
    
    Sheets("勤怠支給控除一覧表").Activate
    Cells(1, 1).Select
    
    MsgBox "処理が完了しました。", Title:=ThisWorkbook.Name
    
Kill:
    Set cc = Nothing
    Set vc = Nothing
    
    Application.DisplayAlerts = True
    
End Sub

'// 不要な列を削除
Private Function deleteUnnecessaryColumns(ByVal companyRow As Long, ByVal headerRow As Long, cc As ChartController) As Boolean

    deleteUnnecessaryColumns = False

    If Sheets("設定").Cells(companyRow, 5).Value = "" Then
        deleteUnnecessaryColumns = True
        Exit Function
    End If

    Dim unnecessaryDepartments As Variant: unnecessaryDepartments = Split(Sheets("設定").Cells(companyRow, 5).Value, "-")
    Dim i As Long
    Dim targetColumn As Long
    
    For i = 0 To UBound(unnecessaryDepartments)
        targetColumn = cc.searchDepartmentColumn(unnecessaryDepartments(i), headerRow)
        If targetColumn = 0 Then: Exit Function
        
        Sheets("勤怠支給控除一覧表").Columns(targetColumn).Delete xlToLeft
    Next
    
    deleteUnnecessaryColumns = True
            
End Function

'// 「事務」にも「現場」にも含まない列を移動
Private Function moveIndependentColumns(ByVal companyRow As Long, ByVal headerRow As Long, cc As ChartController) As Boolean

    moveIndependentColumns = False
    
    With Sheets("設定")
    
        If .Cells(companyRow, 4).Value = "" Then
            moveIndependentColumns = True
            Exit Function
        End If
    
        '// 「事務」にも「現場」にも含まない列
        Dim independentColumns As Variant: independentColumns = Split(.Cells(companyRow, 4).Value, "-")
    End With
    
    Dim i As Long
    Dim targetColumn As Long
    
    With Sheets("勤怠支給控除一覧表")
        .Activate
    
        For i = 0 To UBound(independentColumns)
            
            If independentColumns(i) = 0 Then: GoTo Continue
            
            targetColumn = cc.searchDepartmentColumn(independentColumns(i), headerRow)
            If targetColumn = 0 Then: Exit Function
            
            '// 列移動
            cc.moveColumn targetColumn, 3
Continue:
        Next
    End With
    
    moveIndependentColumns = True
            
End Function

'// 事務列作成
Private Function createOfficeColumn(ByVal companyRow As Long, ByVal headerRow As Long, cc As ChartController, vc As ValueController) As Boolean

    createOfficeColumn = False
    
    '// 事務列を判別するためのランダムのコードをリセット
    Sheets("mode").Cells(1, 2).Value = ""
    
    With Sheets("設定")
    '// 事務列に含む部署が無ければ抜ける
        If .Cells(companyRow, 3).Value = "" Then
            createOfficeColumn = True
            Exit Function
        End If
        
        '// 「事務」に含む部署コード
        Dim officeDepartments As Variant: officeDepartments = Split(.Cells(companyRow, 3).Value, "-")
    End With
            
    With Sheets("勤怠支給控除一覧表")
        .Columns(3).Insert xlToRight
        Dim officeCode As Long: officeCode = vc.generate8DigitsNumber
        .Cells(headerRow, 3).Value = officeCode & " 事務"
        Sheets("mode").Cells(1, 2).Value = officeCode
        
        '//「事務」に含む部署コードをが入力された列番号を格納した配列を作成
        Dim targetColumns As Variant: targetColumns = departmentCodes2ColumnNumbers(officeDepartments, headerRow, cc)
        If targetColumns(0) = "false" Then: Exit Function
        
        '// 人数計算
        Dim totalAmount As Long: totalAmount = sumPeople(targetColumns, headerRow, cc)
        If totalAmount = -1 Then: Exit Function
        .Cells(headerRow + 1, 3).Value = "【 計 " & totalAmount & "名 】"
        
        '// 通勤手当金合計
        Dim sumFormula As String: sumFormula = sumAmountOfMoney(targetColumns, headerRow, headerRow + 2, cc)
        If sumFormula = "false" Then: Exit Function
        .Cells(headerRow + 2, 3).Formula = sumFormula
        
        '// 他の金額も合計が出るようにAutofill
        .Cells(headerRow + 2, 3).AutoFill .Range(.Cells(headerRow + 2, 3), .Cells(.Cells(Rows.Count, 1).End(xlUp).Row, 3))
        .Columns(3).Copy
        .Columns(3).PasteSpecial xlPasteValues
        
        '// 「事務」に含む部署のヘッダーの色変更
        Dim i As Long
        For i = 0 To UBound(targetColumns)
            Cells(headerRow, Int(targetColumns(i))).Font.Color = RGB(51, 153, 51)
        Next
        
    End With

    createOfficeColumn = True

End Function

'// 現場列作成
Private Function createFieldWorkColumn(ByVal companyRow As Long, ByVal headerRow As Long, cc As ChartController) As Boolean

    createFieldWorkColumn = False

    If Sheets("設定").Cells(companyRow, 6).Value = False Then
        createFieldWorkColumn = True
        Exit Function
    End If
    
    With Sheets("勤怠支給控除一覧表")
        .Columns(3).Insert xlToRight
        .Cells(headerRow, 3).Value = "現場"
        
        '/**
         '* 人数計算
        '**/
        
        '// 現場以外の部署コード(「事務」にも「現場」にも含めない部署と事務
        Dim officeCode As String: officeCode = ""
        If Sheets("mode").Cells(1, 2).Value <> "" Then
            officeCode = "-" & Sheets("mode").Cells(1, 2).Value
        End If
        
        Dim notFieldWorkDepartments As Variant: notFieldWorkDepartments = Split(Sheets("設定").Cells(companyRow, 4).Value & officeCode, "-")
        
        '// 現場以外の部署コードをコードが入力されている列番号を格納した配列に変換
        Dim notFieldWorkColumns As Variant: notFieldWorkColumns = departmentCodes2ColumnNumbers(notFieldWorkDepartments, headerRow, cc)
        If notFieldWorkColumns(0) = "false" Then: Exit Function
        
        '//合計から現場以外の人数を引いて現場の人数を求める
        Dim vc As New ValueController
        Dim numberOfFieldWorkers As Long
        
        numberOfFieldWorkers = cc.countPeople(2, headerRow, vc) - sumPeople(notFieldWorkColumns, headerRow, cc)
        .Cells(headerRow + 1, 3).Value = "【 計 " & numberOfFieldWorkers & "名 】"
        
        Set vc = Nothing
        
        '/**
         '* 各種金額入力
        '**/
          
        '// 通勤手当計算
        Dim fieldWorkFormula As String: fieldWorkFormula = "=B" & headerRow + 2
        Dim i As Long
        Dim columnAlphabet As String
        
        For i = 0 To UBound(notFieldWorkColumns)
            columnAlphabet = vc.columnNumber2Alphabet(Int(notFieldWorkColumns(i)))
            fieldWorkFormula = fieldWorkFormula & "-" & columnAlphabet & headerRow + 2
        Next
        
        .Cells(headerRow + 2, 3).Formula = fieldWorkFormula
        
        '// 各種金額にAutoFillで反映させ、数値化
        .Cells(headerRow + 2, 3).AutoFill .Range(.Cells(headerRow + 2, 3), .Cells(.Cells(Rows.Count, 1).End(xlUp).Row, 3))
        .Columns(3).Copy
        .Columns(3).PasteSpecial xlPasteValues
        
        '// 事務列のヘッダーのコード削除
        If Sheets("設定").Cells(companyRow, 3).Value <> "" Then
            .Cells(headerRow, Int(notFieldWorkColumns(UBound(notFieldWorkColumns)))).Value = "事務"
        End If
        
    End With
    
    createFieldWorkColumn = True
    
End Function

'/**
 '* 複数の部署コードが格納された配列から部署コードが入力された列番号を格納した配列を返す
' **/
Private Function departmentCodes2ColumnNumbers(departmentCodes As Variant, ByVal headerRow As Long, cc As ChartController) As Variant

    departmentCodes2ColumnNumbers = Split("false", ",")

    Dim i As Long
    Dim tmpColumn As Long
    Dim targetColumns As String
    
    For i = 0 To UBound(departmentCodes)
        If departmentCodes(i) = "" Then: GoTo Continue
        
        tmpColumn = cc.searchDepartmentColumn(departmentCodes(i), headerRow)
        If tmpColumn = 0 Then: Exit Function
        
        If targetColumns = "" Then
            targetColumns = tmpColumn
        Else
            targetColumns = targetColumns & "," & tmpColumn
        End If
Continue:
    Next
        
    departmentCodes2ColumnNumbers = Split(targetColumns, ",")

End Function

'/**
 '* 総支給から通勤手当を引いた金額を表示する「支給金額」行を作成
'**/
Private Sub createBasicSalaryRow(ByVal headerRow As Long)

    With Sheets("勤怠支給控除一覧表")
        .Rows(headerRow + 3).Insert xlDown
        .Cells(headerRow + 3, 1).Value = "支給金額"
        
        .Cells(headerRow + 3, 2).Formula = "=B" & headerRow + 4 & "-B" & headerRow + 2
        .Cells(headerRow + 3, 2).AutoFill .Range(.Cells(headerRow + 3, 2), Cells(headerRow + 3, .Cells(headerRow, Columns.Count).End(xlToLeft).Column))
    
        .Range(.Cells(headerRow + 4, 1), Cells(headerRow + 4, .Cells(headerRow, Columns.Count).End(xlToLeft).Column)).Borders(xlEdgeBottom).LineStyle = xlDouble
    
    End With

End Sub

'/**
' * 指定の部署の同じ行の金額を足し算するエクセルの式を返す
' * @params targetDepartments 金額をたす部署コード
'**/
Private Function sumAmountOfMoney(targetColumns As Variant, headerRow As Long, ByVal targetRow As Long, cc As ChartController) As String

    Dim i As Long
    Dim vc As New ValueController
    
    Dim returnFormula As String: returnFormula = "=SUM("
    
    '// 列番号をアルファベットに変換したもの
    Dim columnAlphabet As String
    
    For i = 0 To UBound(targetColumns)
        columnAlphabet = vc.columnNumber2Alphabet(Int(targetColumns(i)))
        
        If returnFormula = "=SUM(" Then
            returnFormula = returnFormula & columnAlphabet & targetRow
        Else
            returnFormula = returnFormula & "," & columnAlphabet & targetRow
        End If
    Next
    
    Set vc = Nothing
    
    sumAmountOfMoney = returnFormula & ")"

End Function

'// 複数部署の合計人数を求める
Private Function sumPeople(targetColumns As Variant, headerRow As Long, cc As ChartController) As Long

    Dim i As Long
    Dim totalAmount As Long
    Dim vc As New ValueController
    
    For i = 0 To UBound(targetColumns)
        totalAmount = totalAmount + cc.countPeople(targetColumns(i), headerRow, vc)
    Next
    
    Set vc = Nothing
    
    sumPeople = totalAmount
    
End Function

'// 「控除」のシートの表加工
Private Sub processDeductionChart(vc As ValueController)

    Dim sumRow As Long
    
    With Sheets("控除")
        .Activate
    
        '// 合計金額が0の項目を削除
        On Error Resume Next
        sumRow = WorksheetFunction.Match("【 合計*", .Columns(1), 0)
        On Error GoTo 0
        
        If sumRow = 0 Then: Exit Sub
        
        Dim headerRow As Long: headerRow = Sheets("控除").Cells(1, 3).End(xlDown).Row
        
        Dim i As Long
        Dim lastColumn As Long: lastColumn = .Cells(headerRow, Columns.Count).End(xlToLeft).Column
        
        For i = 4 To lastColumn
            If lastColumn < i Then: Exit For
            
            If .Cells(sumRow, i).Value = 0 Then
                Columns(i).Delete xlToLeft
                i = i - 1
                lastColumn = lastColumn - 1
            End If
        Next
        
        '// 全項目の合計が0円の人を削除
        lastColumn = .Cells(headerRow, Columns.Count).End(xlToLeft).Column
        Dim lastRow As Long: lastRow = .Cells(Rows.Count, 2).End(xlUp).Row
        
        .Cells(headerRow + 1, lastColumn + 1).Formula = "=SUM(D" & headerRow + 1 & ":" & vc.columnNumber2Alphabet(lastColumn) & headerRow + 1 & ")"
        .Cells(headerRow + 1, lastColumn + 1).AutoFill .Range(.Cells(headerRow + 1, lastColumn + 1), .Cells(lastRow, lastColumn + 1))
    
        .Range(.Cells(headerRow, 1), .Cells(lastRow, lastColumn + 1)).AutoFilter lastColumn + 1, "0"
        
        If .Cells(Rows.Count, 2).End(xlUp).Row > headerRow Then
            Dim targetRange As Range: Set targetRange = .Cells(headerRow, lastColumn).CurrentRegion
            targetRange.Offset(headerRow).Resize(targetRange.Rows.Count - headerRow).Delete
            Set targetRange = Nothing
        End If
        
        .Cells(1, 1).AutoFilter
        .Columns(lastColumn + 1).Clear
        .Cells(1, 1).Select
          
    End With
         
End Sub

'// 途中でエラーになった時に表を最初の状態に戻す
Private Sub resetChart()

    Sheets("tmp").Cells.Copy Sheets("勤怠支給控除一覧表").Cells(1, 1)

End Sub

'// 表貼り付けのためのフォーム起動
Public Sub openFormToPasteChart()

    Sheets("mode").Cells(1, 1).Value = "PASTE_CHART"
    frmCompany.Show
    
End Sub

'// 貼り付け
Public Sub pasteChart(ByVal company As String)

    Application.ScreenUpdating = False

    Dim path As String
    
    With Sheets("設定")
        path = .Cells(WorksheetFunction.Match(company, .Columns(1), 0), 2).Value
    End With
    
    If path = "" Then
        MsgBox "貼り付け先が設定されていません。", Title:=ThisWorkbook.Name
    End If
    
    Dim fso As New FileSystemObject
    
    If fso.FileExists(path) = False Then
        MsgBox "貼り付け先ファイルが見つかりませんでした。" & vbLf & "設定を変更してください。" & vbLf & vbLf & "【現在設定中の貼り付け先】" & vbLf & path, vbQuestion, ThisWorkbook.Name
        GoTo Kill
    End If
    
    Workbooks.Open path
    
    Dim fileName As Variant: fileName = Split(path, "\")
    fileName = fileName(UBound(fileName))
    
    ThisWorkbook.Sheets("勤怠支給控除一覧表").Cells.Copy Workbooks(fileName).Sheets("勤怠支給控除一覧表").Cells(1, 1)
    ThisWorkbook.Sheets("控除").Cells.Copy Workbooks(fileName).Sheets("控除").Cells(1, 1)
    
    MsgBox "貼り付けが完了しました。", Title:=ThisWorkbook.Name
    
    ThisWorkbook.Close True

Kill:
    Set fso = Nothing
    
End Sub

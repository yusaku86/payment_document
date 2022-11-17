Attribute VB_Name = "income"
'/**
 '* 給料資料作成
'**/
Option Explicit
'/**
 '* 仕訳追加のためのフォーム起動
 '* @params targetSheetName 仕訳を追加するシートの名前
Public Sub inputData(ByVal targetSheetName As String)

    data.Show vbModeless
    
    data.hiddenTxtTargetSheetName.value = targetSheetName
    
End Sub

'/**
' * YM修理代金入力
' * @params creditColumn 貸方の列番号
' * @params companyName  会社名
' * @params addCheckBox  チェックボックスを追加するか
'**/
Sub inputYM(ByVal creditColumn As Long, ByVal companyName As String, ByVal needCheckBox As Boolean)
    
    Dim msgAns As VbMsgBoxResult: msgAns = MsgBox("YMﾜｰｸｽ修理代を入力します。" & vbLf & "実行してよろしいですか?", vbYesNo + vbQuestion, "給料資料作成:" & companyName)
    If msgAns = vbNo Then
        Exit Sub
    End If
    
    Dim overWrite As Boolean: overWrite = False
    
    '// 既に入力されている場合、上書きするか確認
    On Error Resume Next
    Dim writtenRow As Long: writtenRow = WorksheetFunction.Match("*1123:整備売掛金 YMﾜｰｸｽ修理代*", Sheets(companyName).Columns(creditColumn), 0)
    On Error GoTo 0
    
    If writtenRow > 0 Then
        msgAns = MsgBox("既にYMﾜｰｸｽ修理代が入力されていますが、上書きしますか?" & vbLf & _
                      "(上書きせず、追加する場合には「いいえ」を選択してください。", vbYesNoCancel + vbQuestion, "給料資料作成:" & companyName)
                      
        If msgAns = vbCancel Then
            Exit Sub
        ElseIf msgAns = vbYes Then
            overWrite = True
        End If
    End If
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    '// シート「控除」のYMﾜｰｸｽ修理代の金額が0より大きい人の名前をキー、金額を値として連想配列を作成
    With Sheets("控除")
        On Error Resume Next
        
        Dim targetColumn As Long: targetColumn = WorksheetFunction.Match("YMﾜｰｸｽ修理代", .Rows(5), 0)
        Dim sumRow As Long: sumRow = WorksheetFunction.Match("*合計*", .Columns(1), 0)
        
        On Error GoTo 0
        
        '// YMの修理代がない、もしくは合計行が見つからなければ抜ける
        If targetColumn = 0 Then
            MsgBox "今月のYMﾜｰｸｽ修理代はありません。", Title:=ThisWorkbook.Name
            Exit Sub
        ElseIf sumRow = 0 Then
            MsgBox "「控除」のシートに合計の行がありません。合計の行は削除しないでください。", vbQuestion, "給料資料作成:" & companyName
            Exit Sub
        End If
        
        Dim i As Long
        Dim lastRow As Long: lastRow = .Cells(Rows.Count, 2).End(xlUp).Row
        Dim ymDictionary As New Dictionary
        
        '// YM修理代の金額が0円より大きい人の名前をキー、金額を値にした連想は配列を作成
        For i = 7 To lastRow
            With .Cells(i, targetColumn)
                If .value > 0 Then
                    ymDictionary.add Sheets("控除").Cells(i, 3).value, .value
                End If
            End With
        Next
    End With
    
    '// 作成した配列のキーと勘定科目名他を貸方摘要に、値を金額のセルに入力し、チェックボックス追加
    With Sheets(companyName)
        Dim chkController As New checkBoxController
        Dim targetRow As Long
        Dim cellValue As String
        
        '// 合計が既に表にあるか
        Dim alreadySum As Boolean: alreadySum = WorksheetFunction.CountIf(Columns(creditColumn), "貸方合計額") = 1
        
        For i = 1 To ymDictionary.Count
            targetRow = 0
            lastRow = .Cells(Rows.Count, creditColumn).End(xlUp).Row
            
            '// セルに入力する値
            cellValue = "1123:整備売掛金 YMﾜｰｸｽ修理代" & vbLf & ymDictionary.Keys(i - 1)
            
            '// 上書きする場合
            If overWrite Then
                On Error Resume Next
                targetRow = WorksheetFunction.Match(cellValue, .Columns(creditColumn), 0)
                On Error GoTo 0
            End If
        
            '// 合計が既に表にあり、合計行より上に空欄があり、上書きではない場合
            If alreadySum And Cells(lastRow, creditColumn).End(xlUp).Offset(1).value = "" And targetRow = 0 Then
                targetRow = Cells(lastRow, creditColumn).End(xlUp).Row + 1
            
            '// 合計が既に表にあり、合計行より上に空欄がなく、上書きではない場合
            '// →合計行の上に1行挿入し、合計行のチェックボックスがあれば名前変更
            ElseIf alreadySum And Cells(lastRow, creditColumn).End(xlUp).Offset(1).value <> "" And targetRow = 0 Then
                targetRow = WorksheetFunction.Match("貸方合計額", Columns(creditColumn), 0)
                Range(Cells(targetRow, creditColumn - 2), Cells(targetRow, creditColumn + 2)).Insert xlDown
                
                If chkController.isExistChk(ActiveSheet, "chk" & targetRow) Then
                    ActiveSheet.Shapes("chk" & targetRow).Name = "chk" & targetRow + 1
                End If
            
            '// 合計行がまだなく、上書きではない場合
            ElseIf targetRow = 0 Then
                targetRow = lastRow + 1
            End If
            
            With .Cells(targetRow, creditColumn)
                .value = cellValue
                .HorizontalAlignment = xlLeft
                .Offset(, 1).value = ymDictionary(ymDictionary.Keys(i - 1))
            End With
            
            '// チェックボックス追加
            If needCheckBox Then
                chkController.add .Cells(targetRow, creditColumn + 2), "chk" & targetRow
            End If
        Next
    End With
    
    Set chkController = Nothing
    
    MsgBox "処理が完了しました。", Title:="給料資料作成:" & companyName

End Sub

'/**
' * 合計を計算して入力
' * @params debitAmountColumn 借方金額列
' * @params creditAmounColumn 貸方金額列
' * @params firstRow          先頭行(ヘッダーの次の行)
' * @params needCheckBox      チェックボックスを追加するか
' * @params showMsg           既に合計が入力されている場合にメッセージを表示するか
'**/
Sub calculateTotalAmount(ByVal debitAmountColumn As Long, ByVal creditAmountColumn As Long, ByVal firstRow As Long, ByVal needCheckBox As Boolean, Optional ByVal showMsg As Boolean = True)
    
    Application.ScreenUpdating = False
    
    '// 借方と貸方の最終行の大きい方が合計を入力する行番号
    Dim lastRow As Long
    If Cells(Rows.Count, debitAmountColumn - 1).End(xlUp).Row >= Cells(Rows.Count, creditAmountColumn - 1).End(xlUp).Row Then
        lastRow = Cells(Rows.Count, debitAmountColumn - 1).End(xlUp).Row
    Else
        lastRow = Cells(Rows.Count, creditAmountColumn - 1).End(xlUp).Row
    End If
    
    Dim targetCell As Range
    
    If Application.WorksheetFunction.CountIf(Columns(creditAmountColumn - 1), "貸方合計額") >= 1 Then
    
        Dim overWrite As VbMsgBoxResult
        
        '// showMsgがFalseの場合はメッセージを表示せず再計算
        If showMsg = False Then
            overWrite = vbYes
        Else
            overWrite = MsgBox("既に合計が入力されていますが、再計算しますか?", vbYesNo + vbQuestion, "給料資料作成:" & ActiveSheet.Name)
        End If
        
        If overWrite = vbNo Then
            Exit Sub
        Else
            Set targetCell = Cells(lastRow, creditAmountColumn - 1)
        End If
    Else
        Set targetCell = Cells(lastRow + 1, creditAmountColumn - 1)
    End If
    
    ' // 合計金額の計算
    Dim totalAmount As Long
    totalAmount = WorksheetFunction.Sum(Range(Cells(firstRow, creditAmountColumn), Cells(targetCell.Row - 1, creditAmountColumn)))
    
    '// 借方の未払費用(給料)の計算
    Cells(firstRow, debitAmountColumn).value = totalAmount - _
        WorksheetFunction.Sum(Range(Cells(firstRow + 1, debitAmountColumn), Cells(targetCell.Row - 1, debitAmountColumn)))
    
    '// 借方・貸方の合計表示
    With targetCell
        .value = "貸方合計額"
        .HorizontalAlignment = xlCenter
        .Offset(0, 1).value = totalAmount
        
        With .Offset(0, -2)
            .Formula = "借方合計額"
            .HorizontalAlignment = xlCenter
        End With
        .Offset(0, -1).value = totalAmount
    End With
    
    '// チェックボックス作成
    If overWrite <> vbYes And needCheckBox Then
        Dim chkController As New checkBoxController
        
        With targetCell
        '// chkContoller.add  [追加するセル],[チェックボックスの名前]
            chkController.add Cells(.Row, .Offset(, 2).Column), "chk" & .Row
            Set chkController = Nothing
        End With
    End If
    
    Set targetCell = Nothing
    
    Range(Columns(debitAmountColumn - 1), Columns(creditAmountColumn)).EntireColumn.AutoFit

End Sub
'/**
' * 入力されたデータをクリア
' * @params debitAmountColumn  借方金額列
' * @params creditAmountColumn 貸方金額列
' * @params headerRow          ヘッダー行
' * @params firstClearDebitRow 借方のクリアを行う最初の行
' * @params startClearRow      クリアを開始する行
' * @params existCheckBox      チェックボックスが存在するか
'**/
Sub clearData(ByVal debitAmountColumn As Long, ByVal creditAmountColumn As Long, ByVal headerRow As Long, ByVal firstClearDebitRow As Long, ByVal startClearRow As Long, Optional ByVal existCheckBox As Boolean = True)
    

    Dim boundaryRow As Long
    boundaryRow = Application.InputBox("何行目のデータからクリアするか入力してください。", "給料資料作成:" & ActiveSheet.Name, startClearRow, Type:=1)
    
    Application.ScreenUpdating = False
    
    If boundaryRow = 0 Then
        Exit Sub
    ElseIf boundaryRow <= 0 Then
        MsgBox "入力された値が無効です。", vbQuestion, "給料資料作成:" & ActiveSheet.Name
        Exit Sub
    End If
    
    Dim lastRow As Long: lastRow = Cells(Rows.Count, creditAmountColumn - 1).End(xlUp).Row
    
    '// 借方のクリア
    Cells(headerRow + 1, debitAmountColumn).ClearContents
    Range(Cells(firstClearDebitRow, debitAmountColumn - 1), Cells(lastRow, debitAmountColumn)).ClearContents
        
    '// チェックボックスのチェック解除
    Dim i As Long
    
    If existCheckBox Then
        For i = ActiveSheet.CheckBoxes.Count To 1 Step -1
            ActiveSheet.CheckBoxes(i).value = False
        Next
    End If
    
    '// 貸方クリア&チェックボックス削除
    If lastRow < boundaryRow Then
        Exit Sub
    End If
    
    Range(Cells(boundaryRow, creditAmountColumn - 1), Cells(lastRow, creditAmountColumn)).ClearContents
    
    If existCheckBox Then
        For i = lastRow To boundaryRow Step -1
            On Error Resume Next
            ActiveSheet.Shapes("chk" & i).Delete
            On Error GoTo 0
        Next
    End If
    
End Sub
'// 指定の範囲の印刷
Private Sub printBill(ByVal targetRange As Range, ByVal msgTitle As String, Optional ByVal printOrientation As Long = xlPortrait)
    
    Application.ScreenUpdating = False
    
    'メッセージを表示して確認
    If MsgBox("印刷します。よろしいですか?", vbYesNo + vbQuestion, msgTitle) = vbNo Then
        Exit Sub
    End If
    
    '印刷設定
    With ActiveSheet.PageSetup
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = 1
        .Orientation = printOrientation
    End With
    
    '印刷
    targetRange.PrintOut
    
End Sub
'/**
' * 完成したファイルを指定のフォルダに保存
' * @params companyFolderPath   保存先フォルダ(会社名まで)
' * @params cutoffDate   締め日
' * @params paymentDate  支払い日
' * @params ompanyName   会社名
 '**/
Public Sub registerFile(ByVal companyFolderPath As String, ByVal cutoffDate As Date, ByVal paymentDate As Date, ByVal companyName As String)
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    Dim fso As New FileSystemObject
    
    '// companyFolderPathで指定したフォルダが無ければ抜ける
    If fso.FolderExists(companyFolderPath) = False Then
        MsgBox "保存先のフォルダが存在しません。" & vbLf & "「設定」シートで保存先フォルダを設定してください。", vbQuestion, "給料資料作成:" & companyName
        GoTo KILL
    End If
    
    '// 締め日から保存先フォルダ名を判定
    '// 例:2022年3月31日 → 2021年4月～3月
    '//    2022年4月30日 → 2022年4月～3月
    Dim folderPath As String
    
    If Month(cutoffDate) >= 1 And Month(cutoffDate) <= 3 Then
        folderPath = companyFolderPath & "\" & Year(cutoffDate) - 1 & "年4月～3月"
    Else
        folderPath = companyFolderPath & "\" & Year(cutoffDate) & "年4月～3月"
    End If
    
    '// 保存先フォルダが無ければ作成
    If fso.FolderExists(folderPath) = False Then
        fso.CreateFolder folderPath
    End If
    
    '// 保存するファイル名
    Dim fileName As String
    fileName = folderPath & "\" & companyName & " " & Year(cutoffDate) & "年" & Month(cutoffDate) & "月分(" & Month(paymentDate) & "月支払).xlsm"
    
    '// 既にファイルがある場合は上書きするか確認
    If fso.FileExists(fileName) Then
        If MsgBox("既にファイルがありますが、上書きしますか?", vbYesNo + vbQuestion, "給料資料作成:" & companyName) = vbNo Then
            GoTo KILL
        End If
    End If
    
    '// 既にファイルがあり、開いていると保存できないため開いていたら閉じる
    Dim wb As Workbook
    For Each wb In Workbooks
        If wb.Name = fso.GetFileName(fileName) Then
            wb.Close
        End If
    Next
    
    ActiveWorkbook.SaveCopyAs fileName
    Application.Calculate
    
    MsgBox "登録が完了しました。", Title:="給料資料作成:" & companyName

KILL:
    Set fso = Nothing

End Sub

'// 保存先フォルダ名設定
Public Sub setFolderPath(companyName As String)

    With Application.FileDialog(msoFileDialogFolderPicker)
        .InitialFileName = "G:"
        .AllowMultiSelect = False
        .Title = "保存先フォルダ選択:" & companyName
        
        If .Show Then
            Sheets("設定").Cells(2, 3).value = .SelectedItems(1)
            
            MsgBox "保存先を変更しました。", Title:="給料資料作成:" & companyName
        End If
    End With
    
End Sub


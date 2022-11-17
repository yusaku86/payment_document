VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} data 
   Caption         =   "データ入力"
   ClientHeight    =   5190
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10500
   OleObjectBlob   =   "data.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "data"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'// 給与仕訳を追加するフォーム
Option Explicit

Dim myValidator As New Validator                '// バリテーションを行うクラス

Dim ac As New ArrayController                   '// 配列を制御するクラス

Dim chk As New checkBoxController               '// チェックボックスを制御するクラス

Dim cmbController As New FormComboBoxController '// ユーザーフォームのコンボボックスを制御するクラス

Dim accountLists As Variant                     '// 勘定科目を格納した配列

Dim previousFocus As String                     '// 1つ前にフォーカスがあったコントロール名

Dim reg As New RegExp

'// フォーム起動時の処理
Private Sub UserForm_Initialize()

    Application.ScreenUpdating = False
    
    Dim myControl As Control
    
    '// エラーメッセージ非表示
    For Each myControl In Me.Controls
        If myControl.Name Like "lblErr*" Then
            myControl.Visible = False
        End If
    Next
    
    '// コンボボックスの値設定
    Call addAllAccounts(Me.cmbDebitCode)
    Call addAllAccounts(Me.cmbCreditCode)
    
    Application.ScreenUpdating = True

End Sub

'// 科目コンボボックスのリストに全科目追加
Private Sub addAllAccounts(ByVal targetCmb As ComboBox)

    Dim i As Long

    With Sheets("設定")
    
        '// 科目コードコンボボックスのリスト追加 「科目コード(科目名)」の形
        For i = 2 To .Cells(Rows.Count, 4).End(xlUp).Row
            cmbDebitCode.AddItem .Cells(i, 4).value & "(" & .Cells(i, 6).value & ")"
            cmbCreditCode.AddItem .Cells(i, 4).value & "(" & .Cells(i, 6).value & ")"
        Next
    
    End With
    
End Sub

'// 給料仕訳の追加(メイン)
Private Sub cmdEnter_Click()

    '// 入力欄が全て空欄だったら抜ける
    If anyOneInputed(myValidator) = False Then
        Exit Sub
    End If

    '// バリデーション
    If validateInputData = False Then
        Exit Sub
    End If

    With Sheets(Me.hiddenTxtTargetSheetName.value)

        '/**
         '* データ入力& チェックボックス追加
        '**/
        Dim debitLastRow As Long: debitLastRow = .Cells(Rows.Count, 1).End(xlUp).Row + 1
        Dim creditLastRow As Long: creditLastRow = .Cells(Rows.Count, 3).End(xlUp).Row + 1
         
         
        '// 合計が既に入力されている場合
        Dim alreadySum As Boolean
        If WorksheetFunction.CountIf(.Columns(1), "借方合計額") = 1 Then: alreadySum = True
        
        '// 表の「借方科目」・「貸方科目」の欄に入力する内容
        Dim debitContent As String
        debitContent = Me.cmbDebitCode.value & ":" & Me.txtDebitContent.value & vbLf & Me.txtDebitCustomer.value
         
        Dim creditContent As String
        creditContent = Me.cmbCreditCode.value & ":" & Me.txtCreditContent.value & vbLf & Me.txtCreditCustomer.value
         
        '// 借方のみ入力する場合
        If myValidator.required(Me.txtDebitAmount.value) And myValidator.required(Me.txtCreditAmount.value) = False Then
         
            '// 合計が既に入力されている場合
            If alreadySum Then
                '// 合計金額の行の上に空白の行がある場合
                If .Cells(debitLastRow - 1, 1).End(xlUp).Offset(1).value = "" Then
                    debitLastRow = .Cells(debitLastRow - 1, 1).End(xlUp).Row + 1
                 
                '// 合計金額の行の上に空白の行がない場合:合計行の上に1行挿入、合計行のチェックボックス名変更
                Else
                    debitLastRow = .Cells(Rows.Count, 1).End(xlUp).Row
                    .Range(.Cells(debitLastRow, 1), .Cells(debitLastRow, 5)).Insert xlDown
                
                    If chk.isExistChk(ActiveSheet, "chk" & debitLastRow) = True Then
                        ActiveSheet.Shapes("chk" & debitLastRow).Name = "chk" & debitLastRow + 1
                    End If
                End If
            End If
         
            .Cells(debitLastRow, 1).value = debitContent
            .Cells(debitLastRow, 2).value = Me.txtDebitAmount.value
            .Cells(debitLastRow, 1).HorizontalAlignment = xlLeft
             
            '// チェックボックスがまだ無ければ追加
            If chk.isExistChk(Sheets(Me.hiddenTxtTargetSheetName.value), "chk" & debitLastRow) = False Then
                chk.add .Cells(debitLastRow, 5), "chk" & debitLastRow
            End If
             
        '// 貸方のみ入力する場合
        ElseIf myValidator.required(Me.txtDebitAmount.value) = False And myValidator.required(Me.txtCreditAmount.value) Then
        
            '// 合計が既に入力されている場合
            If alreadySum Then
                '// 合計金額の行の上に空白の行がある場合
                If .Cells(creditLastRow - 1, 3).End(xlUp).Offset(1).value = "" Then
                    creditLastRow = .Cells(creditLastRow - 1, 3).End(xlUp).Row + 1
                    
                '// 合計金額の行の上に空白の行がない場合:合計行の上に1行挿入、合計行のチェックボックス名変更
                Else
                    creditLastRow = .Cells(Rows.Count, 1).End(xlUp).Row
                    Range(.Cells(creditLastRow, 1), .Cells(creditLastRow, 5)).Insert xlDown
                 
                    If chk.isExistChk(ActiveSheet, "chk" & creditLastRow) = True Then
                        ActiveSheet.Shapes("chk" & creditLastRow).Name = "chk" & creditLastRow + 1
                    End If
                End If
            End If
         
            .Cells(creditLastRow, 3).value = creditContent
            .Cells(creditLastRow, 4).value = Me.txtCreditAmount.value
            .Cells(creditLastRow, 3).HorizontalAlignment = xlLeft
             
            '// チェックボックスがまだ無ければ追加
            If chk.isExistChk(ActiveSheet, "chk" & creditLastRow) = False Then
                chk.add .Cells(creditLastRow, 5), "chk" & creditLastRow
            End If
             
        '// 借方・貸方ともに入力する場合
        Else
            Dim lastRow As Long
             
            '// 合計が既に入力されている場合
            If alreadySum Then
                '// 合計金額の行の上に空白の行がない場合:合計行の上に1行挿入、合計行のチェックボックス名変更
                lastRow = Cells(Rows.Count, 1).End(xlUp).Row
                .Range(.Cells(lastRow, 1), .Cells(lastRow, 5)).Insert xlDown
             
                If chk.isExistChk(ActiveSheet, "chk" & lastRow) = True Then
                    ActiveSheet.Shapes("chk" & lastRow).Name = "chk" & lastRow + 1
                End If
         
            ElseIf debitLastRow >= creditLastRow Then
                lastRow = debitLastRow
            Else
                lastRow = creditLastRow
            End If
             
            .Cells(lastRow, 1).value = debitContent
            .Cells(lastRow, 2).value = Me.txtDebitAmount.value
            .Cells(lastRow, 3).value = creditContent
            .Cells(lastRow, 4).value = Me.txtCreditAmount.value
             
            .Cells(lastRow, 1).HorizontalAlignment = xlLeft
            .Cells(lastRow, 3).HorizontalAlignment = xlLeft
             
            chk.add .Cells(lastRow, 5), "chk" & lastRow
        End If
     
    '// 合計が入力されている場合は再計算
        '// calculateTotalAmount([借方金額列],[貸方金額列],[先頭行(ヘッダーの次)],[チェックボックスを追加するか],[再計算のメッセージを表示するか])
        If alreadySum Then
            Call calculateTotalAmount(2, 4, 4, False, False)
        End If
     
        '// テキストボックスの文字初期化
        Call clearInput
     
     
        '// ドロップダウン閉じるために1度他のコントロールにフォーカスを充てる
        Me.cmdClose.SetFocus
        Me.cmbDebitCode.SetFocus
     
        .Range(.Columns(1), .Columns(4)).EntireColumn.AutoFit
        .Cells.EntireRow.AutoFit
    
    
    End With
    
End Sub

'/**
 '* 入力内容のバリデーション
 '* バリデーションに引っかかったら該当のエラーメッセージを表示する
'**/
Private Function validateInputData() As Boolean

    validateInputData = True

    '// 借方の要素の配列
    Dim arrDebit As Variant
    arrDebit = Array(Me.cmbDebitCode, Me.txtDebitContent, Me.txtDebitCustomer, Me.txtDebitAmount)
    
    '// 貸方の要素の配列
    Dim arrCredit As Variant
    arrCredit = Array(Me.cmbCreditCode, Me.txtCreditContent, Me.txtCreditCustomer, Me.txtCreditAmount)

    '/**
     '* 入力の確認
    '**/
    '// requiredWith([配列],値)
    '// 配列のいずれかの値が空白でない場合に、値が空白だとFalseを返す

    '// arrayRemoveIndex([対象の配列],[配列から削除する要素のインデックス])


    '*****借方(借方の要素のうちどれか1つでも入力されている場合)*****
    
    '// 科目コードが入力されているか
    If myValidator.requiredWith(ac.arrayRemoveIndex(arrDebit, 0), Me.cmbDebitCode) = False Then
        Me.lblErrDebitCode.Caption = "入力が必要です。"
        Me.lblErrDebitCode.Visible = True
        validateInputData = False
    Else
        Me.lblErrDebitCode.Visible = False
    End If
    '// 摘要が入力されているか
    If myValidator.requiredWith(ac.arrayRemoveIndex(arrDebit, 2), Me.txtDebitContent) = False Then
        Me.lblErrDebitContent.Caption = "入力が必要です。"
        Me.lblErrDebitContent.Visible = True
        validateInputData = False
    Else
        Me.lblErrDebitContent.Visible = False
    End If
    '// 取引先名が入力されているか
    If myValidator.requiredWith(ac.arrayRemoveIndex(arrDebit, 3), Me.txtDebitCustomer) = False Then
        Me.lblErrDebitCustomer.Caption = "入力が必要です。"
        Me.lblErrDebitCustomer.Visible = True
        validateInputData = False
    Else
        Me.lblErrDebitCustomer.Visible = False
    End If
    '// 金額が入力されているか
    If myValidator.requiredWith(ac.arrayRemoveIndex(arrDebit, 4), Me.txtDebitAmount) = False Then
        Me.lblErrDebitAmount.Caption = "入力が必要です。"
        Me.lblErrDebitAmount.Visible = True
        validateInputData = False
    '// 金額が数字かどうか
    ElseIf myValidator.required(Me.txtDebitAmount) And IsNumeric(Me.txtDebitAmount) = False Then
        Me.lblErrDebitAmount.Caption = "金額には数字を入力してください。"
        Me.lblErrDebitAmount.Visible = True
        validateInputData = False
    Else
        Me.lblErrDebitAmount.Visible = False
    End If
    '***貸方(貸方の要素のうちどれか1つでも入力されている場合)***
    
    '// 科目コードが入力されているか
    If myValidator.requiredWith(ac.arrayRemoveIndex(arrCredit, 0), Me.cmbCreditCode) = False Then
        Me.lblErrCreditCode.Caption = "入力が必要です。"
        Me.lblErrCreditCode.Visible = True
        validateInputData = False
    Else
        Me.lblErrCreditCode.Visible = False
    End If
    '// 摘要が入力されているか
    If myValidator.requiredWith(ac.arrayRemoveIndex(arrCredit, 2), Me.txtCreditContent) = False Then
        Me.lblErrCreditContent.Caption = "入力が必要です。"
        Me.lblErrCreditContent.Visible = True
        validateInputData = False
    Else
        Me.lblErrCreditContent.Visible = False
    End If
    '// 取引先が入力されているか
    If myValidator.requiredWith(ac.arrayRemoveIndex(arrCredit, 3), Me.txtCreditCustomer) = False Then
        Me.lblErrCreditCustomer.Caption = "入力が必要です。"
        Me.lblErrCreditCustomer.Visible = True
        validateInputData = False
    Else
        Me.lblErrCreditCustomer.Visible = False
    End If
    '// 金額が入力されているか
    If myValidator.requiredWith(ac.arrayRemoveIndex(arrCredit, 4), Me.txtCreditAmount) = False Then
        Me.lblErrCreditAmount.Caption = "入力が必要です。"
        Me.lblErrCreditAmount.Visible = True
        validateInputData = False
    '// 金額が数字かどうか
    ElseIf myValidator.required(Me.txtCreditAmount) And IsNumeric(Me.txtCreditAmount) = False Then
        Me.lblErrCreditAmount.Caption = "金額には数字を入力してください。"
        Me.lblErrCreditAmount.Visible = True
        validateInputData = False
    Else
        Me.lblErrCreditAmount.Visible = False
    End If
    
End Function

'// フォームの値のうちどれか1つでも入力されているか確認
Private Function anyOneInputed(myValidator As Validator) As Boolean
    
    Dim myControl As Control
    
    For Each myControl In Me.Controls
    
        '// テキストボックスとコンボボックス以外は跳ばす
        If Not myControl.Name Like "txt*" And Not myControl.Name Like "cmb*" Then
            GoTo Continue
        End If
        
        If myValidator.required(myControl.value) Then
            anyOneInputed = True
            Exit Function
        End If
Continue:
    Next
    
    anyOneInputed = False
            
End Function

'// 入力内容のクリア
Private Sub clearInput()

    Dim myControl As Control
    
    For Each myControl In Me.Controls
        If myControl.Name Like "txt*" Or myControl.Name Like "cmb*" Then
            myControl.value = ""
        End If
    Next

End Sub

'// 閉じるボタンを押したとき
Private Sub cmdClose_Click()

    If anyOneInputed(myValidator) = True Then
        If MsgBox("終了してよろしいですか?", vbYesNo + vbQuestion, "給与仕訳追加") = vbNo Then
            Exit Sub
        End If
    End If

    Unload Me

End Sub

'*************************以下科目コンボボックスの処理******************************

'// 貸方科目の値が変更された時
Private Sub cmbDebitCode_Change()

    Call cmbCodeChanged(cmbDebitCode)
    
End Sub

'// 借方科目コードコンボボックスの値が変更された時
Private Sub cmbCreditCode_Change()

    Call cmbCodeChanged(cmbCreditCode)

End Sub

'/**
 '* 勘定科目コンボボックスの値が変更された時の処理
'**/
Private Sub cmbCodeChanged(ByVal targetCmb As ComboBox)

    '// ドロップダウンを閉じるために他のコントロールにフォーカスする
    Me.cmdClose.SetFocus
    
    '// 入力された値が空白の場合→全科目をリストに追加
    If myValidator.required(targetCmb.value) = False Then
        targetCmb.Clear
        Call addAllAccounts(targetCmb)
        
    '// 値が手入力された場合→入力された値から科目を検索してリストを変更
    ElseIf myValidator.pregMatch(targetCmb.value, "^\d{1,}\(.*\)$", reg) = False Then
        Call clearList(targetCmb)
        Call updateAccountCmbLists(targetCmb)
    End If
        
    targetCmb.SetFocus
    targetCmb.DropDown
    
End Sub

'/**
 '* 勘定科目コンボボックスに入力された値からリストを検索して更新
 '* @params accountCmb 値が変更された勘定科目コンボボックス
'**/
 Private Sub updateAccountCmbLists(ByVal accountCmb As ComboBox)

    Dim i As Long
        
   '// カナ(半角)から勘定科目を検索し、リストに追加する
    Dim inputedKana As String: inputedKana = StrConv(Application.GetPhonetic(accountCmb.value), vbNarrow)

    For i = 2 To Sheets("設定").Cells(Rows.Count, 4).End(xlUp).Row
        If Sheets("設定").Cells(i, 5).value Like inputedKana & "*" Then
            accountCmb.AddItem Sheets("設定").Cells(i, 4).value & "(" & Sheets("設定").Cells(i, 6).value & ")"
        End If
    Next

 End Sub

'// コンボボックスの値を保持してリストのみクリア(リストから選択した値を削除しようとするとエラーがおきるためコンボボックスに手入力した場合のみ使用)
Private Sub clearList(cmb As ComboBox)

    Dim i As Long
    
    For i = cmb.ListCount - 1 To 0 Step -1
        cmb.RemoveItem (i)
    Next
    
End Sub

'************************************借方・貸方の値をコピーする機能***************************************

'// 借方摘要テキストボックスにフォーカスが当たった時
Private Sub txtDebitContent_Enter()

    '// 借方科目コンボボックスからフォーカスが移動した場合→貸方摘要の内容をコピー
    If previousFocus = "cmbDebitCode" Then
        Me.txtDebitContent.value = Me.txtCreditContent
    End If

End Sub

'// 借方取引先テキストボックスにフォーカスが当たった時
Private Sub txtDebitCustomer_Enter()

    '// 借方摘要テキストボックスからフォーカスが移動した場合→貸方取引先をコピー
    If previousFocus = "txtDebitContent" Then
        Me.txtDebitCustomer.value = Me.txtCreditCustomer.value
    End If

End Sub

'// 貸方科目コンボボックスにフォーカスが当たった時
Private Sub cmbCreditCode_Enter()

    '// 借方金額テキストボックス以外からフォーカスが当たった時は抜ける
    If previousFocus <> "txtDebitAmount" Then: Exit Sub

    '// 借方に何かしらデータが入力されており、借方金額が空欄の場合→貸方金額をコピー
    If myValidator.requiredWith(Array(cmbDebitCode.value, txtDebitContent.value, txtDebitCustomer.value), txtDebitAmount.value) = False Then
        Me.txtDebitAmount.value = Me.txtCreditAmount.value
    End If
    
End Sub

'// 貸方摘要テキストボックスにフォーカスが当たった時
Private Sub txtCreditContent_Enter()

    '// 貸方科目コンボボックスからフォーカスが移動した場合→借方摘要の内容をコピー
    If previousFocus = "cmbCreditCode" Then
        Me.txtCreditContent.value = Me.txtDebitContent.value
    End If

End Sub

'// 貸方取引先テキストボックスにフォーカスが当たった時
Private Sub txtCreditCustomer_Enter()

    '// 貸方摘要テキストボックスからフォーカスが移動した場合→借方取引先をコピー
    If previousFocus = "txtCreditContent" Then
        Me.txtCreditCustomer.value = Me.txtDebitCustomer.value
    End If

End Sub

'// 「登録」ボタンにフォーカスが当たった時
Private Sub cmdEnter_Enter()

    '// 貸方金額テキストボックス以外からフォーカスが移動した場合は抜ける
    If previousFocus <> "txtCreditAmount" Then: Exit Sub
    
    '// 貸方に何かしらデータが入力されており、貸方金額が空欄の場合→借方金額をコピー
    If myValidator.requiredWith(Array(cmbCreditCode.value, txtCreditContent.value, txtCreditCustomer.value), txtCreditAmount.value) = False Then
        Me.txtCreditAmount.value = Me.txtDebitAmount.value
    End If

End Sub

'*************************************各コントロールからフォーカスが外れた時の処理************************

Private Sub cmbDebitCode_Exit(ByVal Cancel As MSForms.ReturnBoolean)

    previousFocus = "cmbDebitCode"

End Sub

Private Sub txtDebitContent_Exit(ByVal Cancel As MSForms.ReturnBoolean)

    previousFocus = "txtDebitContent"

End Sub

Private Sub txtDebitCustomer_Exit(ByVal Cancel As MSForms.ReturnBoolean)

    previousFocus = "txtDebitCustomer"

End Sub

Private Sub txtDebitAmount_Exit(ByVal Cancel As MSForms.ReturnBoolean)

    previousFocus = "txtDebitAmount"

End Sub

Private Sub cmbCreditCode_Exit(ByVal Cancel As MSForms.ReturnBoolean)

    previousFocus = "cmbCreditCode"

End Sub

Private Sub txtCreditContent_Exit(ByVal Cancel As MSForms.ReturnBoolean)

    previousFocus = "txtCreditContent"

End Sub

Private Sub txtCreditCustomer_Exit(ByVal Cancel As MSForms.ReturnBoolean)

    previousFocus = "txtCreditCustomer"

End Sub

Private Sub txtCreditAmount_Exit(ByVal Cancel As MSForms.ReturnBoolean)

    previousFocus = "txtCreditAmount"

End Sub

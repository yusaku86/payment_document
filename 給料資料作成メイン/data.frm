VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} data 
   Caption         =   "�f�[�^����"
   ClientHeight    =   5190
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10500
   OleObjectBlob   =   "data.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "data"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'// ���^�d���ǉ�����t�H�[��
Option Explicit

Dim myValidator As New Validator                '// �o���e�[�V�������s���N���X

Dim ac As New ArrayController                   '// �z��𐧌䂷��N���X

Dim chk As New checkBoxController               '// �`�F�b�N�{�b�N�X�𐧌䂷��N���X

Dim cmbController As New FormComboBoxController '// ���[�U�[�t�H�[���̃R���{�{�b�N�X�𐧌䂷��N���X

Dim accountLists As Variant                     '// ����Ȗڂ��i�[�����z��

Dim previousFocus As String                     '// 1�O�Ƀt�H�[�J�X���������R���g���[����

Dim reg As New RegExp

'// �t�H�[���N�����̏���
Private Sub UserForm_Initialize()

    Application.ScreenUpdating = False
    
    Dim myControl As Control
    
    '// �G���[���b�Z�[�W��\��
    For Each myControl In Me.Controls
        If myControl.Name Like "lblErr*" Then
            myControl.Visible = False
        End If
    Next
    
    '// �R���{�{�b�N�X�̒l�ݒ�
    Call addAllAccounts(Me.cmbDebitCode)
    Call addAllAccounts(Me.cmbCreditCode)
    
    Application.ScreenUpdating = True

End Sub

'// �ȖڃR���{�{�b�N�X�̃��X�g�ɑS�Ȗڒǉ�
Private Sub addAllAccounts(ByVal targetCmb As ComboBox)

    Dim i As Long

    With Sheets("�ݒ�")
    
        '// �ȖڃR�[�h�R���{�{�b�N�X�̃��X�g�ǉ� �u�ȖڃR�[�h(�Ȗږ�)�v�̌`
        For i = 2 To .Cells(Rows.Count, 4).End(xlUp).Row
            cmbDebitCode.AddItem .Cells(i, 4).value & "(" & .Cells(i, 6).value & ")"
            cmbCreditCode.AddItem .Cells(i, 4).value & "(" & .Cells(i, 6).value & ")"
        Next
    
    End With
    
End Sub

'// �����d��̒ǉ�(���C��)
Private Sub cmdEnter_Click()

    '// ���͗����S�ċ󗓂������甲����
    If anyOneInputed(myValidator) = False Then
        Exit Sub
    End If

    '// �o���f�[�V����
    If validateInputData = False Then
        Exit Sub
    End If

    With Sheets(Me.hiddenTxtTargetSheetName.value)

        '/**
         '* �f�[�^����& �`�F�b�N�{�b�N�X�ǉ�
        '**/
        Dim debitLastRow As Long: debitLastRow = .Cells(Rows.Count, 1).End(xlUp).Row + 1
        Dim creditLastRow As Long: creditLastRow = .Cells(Rows.Count, 3).End(xlUp).Row + 1
         
         
        '// ���v�����ɓ��͂���Ă���ꍇ
        Dim alreadySum As Boolean
        If WorksheetFunction.CountIf(.Columns(1), "�ؕ����v�z") = 1 Then: alreadySum = True
        
        '// �\�́u�ؕ��Ȗځv�E�u�ݕ��Ȗځv�̗��ɓ��͂�����e
        Dim debitContent As String
        debitContent = Me.cmbDebitCode.value & ":" & Me.txtDebitContent.value & vbLf & Me.txtDebitCustomer.value
         
        Dim creditContent As String
        creditContent = Me.cmbCreditCode.value & ":" & Me.txtCreditContent.value & vbLf & Me.txtCreditCustomer.value
         
        '// �ؕ��̂ݓ��͂���ꍇ
        If myValidator.required(Me.txtDebitAmount.value) And myValidator.required(Me.txtCreditAmount.value) = False Then
         
            '// ���v�����ɓ��͂���Ă���ꍇ
            If alreadySum Then
                '// ���v���z�̍s�̏�ɋ󔒂̍s������ꍇ
                If .Cells(debitLastRow - 1, 1).End(xlUp).Offset(1).value = "" Then
                    debitLastRow = .Cells(debitLastRow - 1, 1).End(xlUp).Row + 1
                 
                '// ���v���z�̍s�̏�ɋ󔒂̍s���Ȃ��ꍇ:���v�s�̏��1�s�}���A���v�s�̃`�F�b�N�{�b�N�X���ύX
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
             
            '// �`�F�b�N�{�b�N�X���܂�������Βǉ�
            If chk.isExistChk(Sheets(Me.hiddenTxtTargetSheetName.value), "chk" & debitLastRow) = False Then
                chk.add .Cells(debitLastRow, 5), "chk" & debitLastRow
            End If
             
        '// �ݕ��̂ݓ��͂���ꍇ
        ElseIf myValidator.required(Me.txtDebitAmount.value) = False And myValidator.required(Me.txtCreditAmount.value) Then
        
            '// ���v�����ɓ��͂���Ă���ꍇ
            If alreadySum Then
                '// ���v���z�̍s�̏�ɋ󔒂̍s������ꍇ
                If .Cells(creditLastRow - 1, 3).End(xlUp).Offset(1).value = "" Then
                    creditLastRow = .Cells(creditLastRow - 1, 3).End(xlUp).Row + 1
                    
                '// ���v���z�̍s�̏�ɋ󔒂̍s���Ȃ��ꍇ:���v�s�̏��1�s�}���A���v�s�̃`�F�b�N�{�b�N�X���ύX
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
             
            '// �`�F�b�N�{�b�N�X���܂�������Βǉ�
            If chk.isExistChk(ActiveSheet, "chk" & creditLastRow) = False Then
                chk.add .Cells(creditLastRow, 5), "chk" & creditLastRow
            End If
             
        '// �ؕ��E�ݕ��Ƃ��ɓ��͂���ꍇ
        Else
            Dim lastRow As Long
             
            '// ���v�����ɓ��͂���Ă���ꍇ
            If alreadySum Then
                '// ���v���z�̍s�̏�ɋ󔒂̍s���Ȃ��ꍇ:���v�s�̏��1�s�}���A���v�s�̃`�F�b�N�{�b�N�X���ύX
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
     
    '// ���v�����͂���Ă���ꍇ�͍Čv�Z
        '// calculateTotalAmount([�ؕ����z��],[�ݕ����z��],[�擪�s(�w�b�_�[�̎�)],[�`�F�b�N�{�b�N�X��ǉ����邩],[�Čv�Z�̃��b�Z�[�W��\�����邩])
        If alreadySum Then
            Call calculateTotalAmount(2, 4, 4, False, False)
        End If
     
        '// �e�L�X�g�{�b�N�X�̕���������
        Call clearInput
     
     
        '// �h���b�v�_�E�����邽�߂�1�x���̃R���g���[���Ƀt�H�[�J�X���[�Ă�
        Me.cmdClose.SetFocus
        Me.cmbDebitCode.SetFocus
     
        .Range(.Columns(1), .Columns(4)).EntireColumn.AutoFit
        .Cells.EntireRow.AutoFit
    
    
    End With
    
End Sub

'/**
 '* ���͓��e�̃o���f�[�V����
 '* �o���f�[�V�����Ɉ�������������Y���̃G���[���b�Z�[�W��\������
'**/
Private Function validateInputData() As Boolean

    validateInputData = True

    '// �ؕ��̗v�f�̔z��
    Dim arrDebit As Variant
    arrDebit = Array(Me.cmbDebitCode, Me.txtDebitContent, Me.txtDebitCustomer, Me.txtDebitAmount)
    
    '// �ݕ��̗v�f�̔z��
    Dim arrCredit As Variant
    arrCredit = Array(Me.cmbCreditCode, Me.txtCreditContent, Me.txtCreditCustomer, Me.txtCreditAmount)

    '/**
     '* ���͂̊m�F
    '**/
    '// requiredWith([�z��],�l)
    '// �z��̂����ꂩ�̒l���󔒂łȂ��ꍇ�ɁA�l���󔒂���False��Ԃ�

    '// arrayRemoveIndex([�Ώۂ̔z��],[�z�񂩂�폜����v�f�̃C���f�b�N�X])


    '*****�ؕ�(�ؕ��̗v�f�̂����ǂꂩ1�ł����͂���Ă���ꍇ)*****
    
    '// �ȖڃR�[�h�����͂���Ă��邩
    If myValidator.requiredWith(ac.arrayRemoveIndex(arrDebit, 0), Me.cmbDebitCode) = False Then
        Me.lblErrDebitCode.Caption = "���͂��K�v�ł��B"
        Me.lblErrDebitCode.Visible = True
        validateInputData = False
    Else
        Me.lblErrDebitCode.Visible = False
    End If
    '// �E�v�����͂���Ă��邩
    If myValidator.requiredWith(ac.arrayRemoveIndex(arrDebit, 2), Me.txtDebitContent) = False Then
        Me.lblErrDebitContent.Caption = "���͂��K�v�ł��B"
        Me.lblErrDebitContent.Visible = True
        validateInputData = False
    Else
        Me.lblErrDebitContent.Visible = False
    End If
    '// ����於�����͂���Ă��邩
    If myValidator.requiredWith(ac.arrayRemoveIndex(arrDebit, 3), Me.txtDebitCustomer) = False Then
        Me.lblErrDebitCustomer.Caption = "���͂��K�v�ł��B"
        Me.lblErrDebitCustomer.Visible = True
        validateInputData = False
    Else
        Me.lblErrDebitCustomer.Visible = False
    End If
    '// ���z�����͂���Ă��邩
    If myValidator.requiredWith(ac.arrayRemoveIndex(arrDebit, 4), Me.txtDebitAmount) = False Then
        Me.lblErrDebitAmount.Caption = "���͂��K�v�ł��B"
        Me.lblErrDebitAmount.Visible = True
        validateInputData = False
    '// ���z���������ǂ���
    ElseIf myValidator.required(Me.txtDebitAmount) And IsNumeric(Me.txtDebitAmount) = False Then
        Me.lblErrDebitAmount.Caption = "���z�ɂ͐�������͂��Ă��������B"
        Me.lblErrDebitAmount.Visible = True
        validateInputData = False
    Else
        Me.lblErrDebitAmount.Visible = False
    End If
    '***�ݕ�(�ݕ��̗v�f�̂����ǂꂩ1�ł����͂���Ă���ꍇ)***
    
    '// �ȖڃR�[�h�����͂���Ă��邩
    If myValidator.requiredWith(ac.arrayRemoveIndex(arrCredit, 0), Me.cmbCreditCode) = False Then
        Me.lblErrCreditCode.Caption = "���͂��K�v�ł��B"
        Me.lblErrCreditCode.Visible = True
        validateInputData = False
    Else
        Me.lblErrCreditCode.Visible = False
    End If
    '// �E�v�����͂���Ă��邩
    If myValidator.requiredWith(ac.arrayRemoveIndex(arrCredit, 2), Me.txtCreditContent) = False Then
        Me.lblErrCreditContent.Caption = "���͂��K�v�ł��B"
        Me.lblErrCreditContent.Visible = True
        validateInputData = False
    Else
        Me.lblErrCreditContent.Visible = False
    End If
    '// ����悪���͂���Ă��邩
    If myValidator.requiredWith(ac.arrayRemoveIndex(arrCredit, 3), Me.txtCreditCustomer) = False Then
        Me.lblErrCreditCustomer.Caption = "���͂��K�v�ł��B"
        Me.lblErrCreditCustomer.Visible = True
        validateInputData = False
    Else
        Me.lblErrCreditCustomer.Visible = False
    End If
    '// ���z�����͂���Ă��邩
    If myValidator.requiredWith(ac.arrayRemoveIndex(arrCredit, 4), Me.txtCreditAmount) = False Then
        Me.lblErrCreditAmount.Caption = "���͂��K�v�ł��B"
        Me.lblErrCreditAmount.Visible = True
        validateInputData = False
    '// ���z���������ǂ���
    ElseIf myValidator.required(Me.txtCreditAmount) And IsNumeric(Me.txtCreditAmount) = False Then
        Me.lblErrCreditAmount.Caption = "���z�ɂ͐�������͂��Ă��������B"
        Me.lblErrCreditAmount.Visible = True
        validateInputData = False
    Else
        Me.lblErrCreditAmount.Visible = False
    End If
    
End Function

'// �t�H�[���̒l�̂����ǂꂩ1�ł����͂���Ă��邩�m�F
Private Function anyOneInputed(myValidator As Validator) As Boolean
    
    Dim myControl As Control
    
    For Each myControl In Me.Controls
    
        '// �e�L�X�g�{�b�N�X�ƃR���{�{�b�N�X�ȊO�͒��΂�
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

'// ���͓��e�̃N���A
Private Sub clearInput()

    Dim myControl As Control
    
    For Each myControl In Me.Controls
        If myControl.Name Like "txt*" Or myControl.Name Like "cmb*" Then
            myControl.value = ""
        End If
    Next

End Sub

'// ����{�^�����������Ƃ�
Private Sub cmdClose_Click()

    If anyOneInputed(myValidator) = True Then
        If MsgBox("�I�����Ă�낵���ł���?", vbYesNo + vbQuestion, "���^�d��ǉ�") = vbNo Then
            Exit Sub
        End If
    End If

    Unload Me

End Sub

'*************************�ȉ��ȖڃR���{�{�b�N�X�̏���******************************

'// �ݕ��Ȗڂ̒l���ύX���ꂽ��
Private Sub cmbDebitCode_Change()

    Call cmbCodeChanged(cmbDebitCode)
    
End Sub

'// �ؕ��ȖڃR�[�h�R���{�{�b�N�X�̒l���ύX���ꂽ��
Private Sub cmbCreditCode_Change()

    Call cmbCodeChanged(cmbCreditCode)

End Sub

'/**
 '* ����ȖڃR���{�{�b�N�X�̒l���ύX���ꂽ���̏���
'**/
Private Sub cmbCodeChanged(ByVal targetCmb As ComboBox)

    '// �h���b�v�_�E������邽�߂ɑ��̃R���g���[���Ƀt�H�[�J�X����
    Me.cmdClose.SetFocus
    
    '// ���͂��ꂽ�l���󔒂̏ꍇ���S�Ȗڂ����X�g�ɒǉ�
    If myValidator.required(targetCmb.value) = False Then
        targetCmb.Clear
        Call addAllAccounts(targetCmb)
        
    '// �l������͂��ꂽ�ꍇ�����͂��ꂽ�l����Ȗڂ��������ă��X�g��ύX
    ElseIf myValidator.pregMatch(targetCmb.value, "^\d{1,}\(.*\)$", reg) = False Then
        Call clearList(targetCmb)
        Call updateAccountCmbLists(targetCmb)
    End If
        
    targetCmb.SetFocus
    targetCmb.DropDown
    
End Sub

'/**
 '* ����ȖڃR���{�{�b�N�X�ɓ��͂��ꂽ�l���烊�X�g���������čX�V
 '* @params accountCmb �l���ύX���ꂽ����ȖڃR���{�{�b�N�X
'**/
 Private Sub updateAccountCmbLists(ByVal accountCmb As ComboBox)

    Dim i As Long
        
   '// �J�i(���p)���犨��Ȗڂ��������A���X�g�ɒǉ�����
    Dim inputedKana As String: inputedKana = StrConv(Application.GetPhonetic(accountCmb.value), vbNarrow)

    For i = 2 To Sheets("�ݒ�").Cells(Rows.Count, 4).End(xlUp).Row
        If Sheets("�ݒ�").Cells(i, 5).value Like inputedKana & "*" Then
            accountCmb.AddItem Sheets("�ݒ�").Cells(i, 4).value & "(" & Sheets("�ݒ�").Cells(i, 6).value & ")"
        End If
    Next

 End Sub

'// �R���{�{�b�N�X�̒l��ێ����ă��X�g�̂݃N���A(���X�g����I�������l���폜���悤�Ƃ���ƃG���[�������邽�߃R���{�{�b�N�X�Ɏ���͂����ꍇ�̂ݎg�p)
Private Sub clearList(cmb As ComboBox)

    Dim i As Long
    
    For i = cmb.ListCount - 1 To 0 Step -1
        cmb.RemoveItem (i)
    Next
    
End Sub

'************************************�ؕ��E�ݕ��̒l���R�s�[����@�\***************************************

'// �ؕ��E�v�e�L�X�g�{�b�N�X�Ƀt�H�[�J�X������������
Private Sub txtDebitContent_Enter()

    '// �ؕ��ȖڃR���{�{�b�N�X����t�H�[�J�X���ړ������ꍇ���ݕ��E�v�̓��e���R�s�[
    If previousFocus = "cmbDebitCode" Then
        Me.txtDebitContent.value = Me.txtCreditContent
    End If

End Sub

'// �ؕ������e�L�X�g�{�b�N�X�Ƀt�H�[�J�X������������
Private Sub txtDebitCustomer_Enter()

    '// �ؕ��E�v�e�L�X�g�{�b�N�X����t�H�[�J�X���ړ������ꍇ���ݕ��������R�s�[
    If previousFocus = "txtDebitContent" Then
        Me.txtDebitCustomer.value = Me.txtCreditCustomer.value
    End If

End Sub

'// �ݕ��ȖڃR���{�{�b�N�X�Ƀt�H�[�J�X������������
Private Sub cmbCreditCode_Enter()

    '// �ؕ����z�e�L�X�g�{�b�N�X�ȊO����t�H�[�J�X�������������͔�����
    If previousFocus <> "txtDebitAmount" Then: Exit Sub

    '// �ؕ��ɉ�������f�[�^�����͂���Ă���A�ؕ����z���󗓂̏ꍇ���ݕ����z���R�s�[
    If myValidator.requiredWith(Array(cmbDebitCode.value, txtDebitContent.value, txtDebitCustomer.value), txtDebitAmount.value) = False Then
        Me.txtDebitAmount.value = Me.txtCreditAmount.value
    End If
    
End Sub

'// �ݕ��E�v�e�L�X�g�{�b�N�X�Ƀt�H�[�J�X������������
Private Sub txtCreditContent_Enter()

    '// �ݕ��ȖڃR���{�{�b�N�X����t�H�[�J�X���ړ������ꍇ���ؕ��E�v�̓��e���R�s�[
    If previousFocus = "cmbCreditCode" Then
        Me.txtCreditContent.value = Me.txtDebitContent.value
    End If

End Sub

'// �ݕ������e�L�X�g�{�b�N�X�Ƀt�H�[�J�X������������
Private Sub txtCreditCustomer_Enter()

    '// �ݕ��E�v�e�L�X�g�{�b�N�X����t�H�[�J�X���ړ������ꍇ���ؕ��������R�s�[
    If previousFocus = "txtCreditContent" Then
        Me.txtCreditCustomer.value = Me.txtDebitCustomer.value
    End If

End Sub

'// �u�o�^�v�{�^���Ƀt�H�[�J�X������������
Private Sub cmdEnter_Enter()

    '// �ݕ����z�e�L�X�g�{�b�N�X�ȊO����t�H�[�J�X���ړ������ꍇ�͔�����
    If previousFocus <> "txtCreditAmount" Then: Exit Sub
    
    '// �ݕ��ɉ�������f�[�^�����͂���Ă���A�ݕ����z���󗓂̏ꍇ���ؕ����z���R�s�[
    If myValidator.requiredWith(Array(cmbCreditCode.value, txtCreditContent.value, txtCreditCustomer.value), txtCreditAmount.value) = False Then
        Me.txtCreditAmount.value = Me.txtDebitAmount.value
    End If

End Sub

'*************************************�e�R���g���[������t�H�[�J�X���O�ꂽ���̏���************************

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

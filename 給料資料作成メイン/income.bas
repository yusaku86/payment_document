Attribute VB_Name = "income"
'/**
 '* ���������쐬
'**/
Option Explicit
'/**
 '* �d��ǉ��̂��߂̃t�H�[���N��
 '* @params targetSheetName �d���ǉ�����V�[�g�̖��O
Public Sub inputData(ByVal targetSheetName As String)

    data.Show vbModeless
    
    data.hiddenTxtTargetSheetName.value = targetSheetName
    
End Sub

'/**
' * YM�C���������
' * @params creditColumn �ݕ��̗�ԍ�
' * @params companyName  ��Ж�
' * @params addCheckBox  �`�F�b�N�{�b�N�X��ǉ����邩
'**/
Sub inputYM(ByVal creditColumn As Long, ByVal companyName As String, ByVal needCheckBox As Boolean)
    
    Dim msgAns As VbMsgBoxResult: msgAns = MsgBox("YMܰ���C�������͂��܂��B" & vbLf & "���s���Ă�낵���ł���?", vbYesNo + vbQuestion, "���������쐬:" & companyName)
    If msgAns = vbNo Then
        Exit Sub
    End If
    
    Dim overWrite As Boolean: overWrite = False
    
    '// ���ɓ��͂���Ă���ꍇ�A�㏑�����邩�m�F
    On Error Resume Next
    Dim writtenRow As Long: writtenRow = WorksheetFunction.Match("*1123:�������|�� YMܰ���C����*", Sheets(companyName).Columns(creditColumn), 0)
    On Error GoTo 0
    
    If writtenRow > 0 Then
        msgAns = MsgBox("����YMܰ���C���オ���͂���Ă��܂����A�㏑�����܂���?" & vbLf & _
                      "(�㏑�������A�ǉ�����ꍇ�ɂ́u�������v��I�����Ă��������B", vbYesNoCancel + vbQuestion, "���������쐬:" & companyName)
                      
        If msgAns = vbCancel Then
            Exit Sub
        ElseIf msgAns = vbYes Then
            overWrite = True
        End If
    End If
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    '// �V�[�g�u�T���v��YMܰ���C����̋��z��0���傫���l�̖��O���L�[�A���z��l�Ƃ��ĘA�z�z����쐬
    With Sheets("�T��")
        On Error Resume Next
        
        Dim targetColumn As Long: targetColumn = WorksheetFunction.Match("YMܰ���C����", .Rows(5), 0)
        Dim sumRow As Long: sumRow = WorksheetFunction.Match("*���v*", .Columns(1), 0)
        
        On Error GoTo 0
        
        '// YM�̏C���オ�Ȃ��A�������͍��v�s��������Ȃ���Δ�����
        If targetColumn = 0 Then
            MsgBox "������YMܰ���C����͂���܂���B", Title:=ThisWorkbook.Name
            Exit Sub
        ElseIf sumRow = 0 Then
            MsgBox "�u�T���v�̃V�[�g�ɍ��v�̍s������܂���B���v�̍s�͍폜���Ȃ��ł��������B", vbQuestion, "���������쐬:" & companyName
            Exit Sub
        End If
        
        Dim i As Long
        Dim lastRow As Long: lastRow = .Cells(Rows.Count, 2).End(xlUp).Row
        Dim ymDictionary As New Dictionary
        
        '// YM�C����̋��z��0�~���傫���l�̖��O���L�[�A���z��l�ɂ����A�z�͔z����쐬
        For i = 7 To lastRow
            With .Cells(i, targetColumn)
                If .value > 0 Then
                    ymDictionary.add Sheets("�T��").Cells(i, 3).value, .value
                End If
            End With
        Next
    End With
    
    '// �쐬�����z��̃L�[�Ɗ���Ȗږ�����ݕ��E�v�ɁA�l�����z�̃Z���ɓ��͂��A�`�F�b�N�{�b�N�X�ǉ�
    With Sheets(companyName)
        Dim chkController As New checkBoxController
        Dim targetRow As Long
        Dim cellValue As String
        
        '// ���v�����ɕ\�ɂ��邩
        Dim alreadySum As Boolean: alreadySum = WorksheetFunction.CountIf(Columns(creditColumn), "�ݕ����v�z") = 1
        
        For i = 1 To ymDictionary.Count
            targetRow = 0
            lastRow = .Cells(Rows.Count, creditColumn).End(xlUp).Row
            
            '// �Z���ɓ��͂���l
            cellValue = "1123:�������|�� YMܰ���C����" & vbLf & ymDictionary.Keys(i - 1)
            
            '// �㏑������ꍇ
            If overWrite Then
                On Error Resume Next
                targetRow = WorksheetFunction.Match(cellValue, .Columns(creditColumn), 0)
                On Error GoTo 0
            End If
        
            '// ���v�����ɕ\�ɂ���A���v�s����ɋ󗓂�����A�㏑���ł͂Ȃ��ꍇ
            If alreadySum And Cells(lastRow, creditColumn).End(xlUp).Offset(1).value = "" And targetRow = 0 Then
                targetRow = Cells(lastRow, creditColumn).End(xlUp).Row + 1
            
            '// ���v�����ɕ\�ɂ���A���v�s����ɋ󗓂��Ȃ��A�㏑���ł͂Ȃ��ꍇ
            '// �����v�s�̏��1�s�}�����A���v�s�̃`�F�b�N�{�b�N�X������Ζ��O�ύX
            ElseIf alreadySum And Cells(lastRow, creditColumn).End(xlUp).Offset(1).value <> "" And targetRow = 0 Then
                targetRow = WorksheetFunction.Match("�ݕ����v�z", Columns(creditColumn), 0)
                Range(Cells(targetRow, creditColumn - 2), Cells(targetRow, creditColumn + 2)).Insert xlDown
                
                If chkController.isExistChk(ActiveSheet, "chk" & targetRow) Then
                    ActiveSheet.Shapes("chk" & targetRow).Name = "chk" & targetRow + 1
                End If
            
            '// ���v�s���܂��Ȃ��A�㏑���ł͂Ȃ��ꍇ
            ElseIf targetRow = 0 Then
                targetRow = lastRow + 1
            End If
            
            With .Cells(targetRow, creditColumn)
                .value = cellValue
                .HorizontalAlignment = xlLeft
                .Offset(, 1).value = ymDictionary(ymDictionary.Keys(i - 1))
            End With
            
            '// �`�F�b�N�{�b�N�X�ǉ�
            If needCheckBox Then
                chkController.add .Cells(targetRow, creditColumn + 2), "chk" & targetRow
            End If
        Next
    End With
    
    Set chkController = Nothing
    
    MsgBox "�������������܂����B", Title:="���������쐬:" & companyName

End Sub

'/**
' * ���v���v�Z���ē���
' * @params debitAmountColumn �ؕ����z��
' * @params creditAmounColumn �ݕ����z��
' * @params firstRow          �擪�s(�w�b�_�[�̎��̍s)
' * @params needCheckBox      �`�F�b�N�{�b�N�X��ǉ����邩
' * @params showMsg           ���ɍ��v�����͂���Ă���ꍇ�Ƀ��b�Z�[�W��\�����邩
'**/
Sub calculateTotalAmount(ByVal debitAmountColumn As Long, ByVal creditAmountColumn As Long, ByVal firstRow As Long, ByVal needCheckBox As Boolean, Optional ByVal showMsg As Boolean = True)
    
    Application.ScreenUpdating = False
    
    '// �ؕ��Ƒݕ��̍ŏI�s�̑傫���������v����͂���s�ԍ�
    Dim lastRow As Long
    If Cells(Rows.Count, debitAmountColumn - 1).End(xlUp).Row >= Cells(Rows.Count, creditAmountColumn - 1).End(xlUp).Row Then
        lastRow = Cells(Rows.Count, debitAmountColumn - 1).End(xlUp).Row
    Else
        lastRow = Cells(Rows.Count, creditAmountColumn - 1).End(xlUp).Row
    End If
    
    Dim targetCell As Range
    
    If Application.WorksheetFunction.CountIf(Columns(creditAmountColumn - 1), "�ݕ����v�z") >= 1 Then
    
        Dim overWrite As VbMsgBoxResult
        
        '// showMsg��False�̏ꍇ�̓��b�Z�[�W��\�������Čv�Z
        If showMsg = False Then
            overWrite = vbYes
        Else
            overWrite = MsgBox("���ɍ��v�����͂���Ă��܂����A�Čv�Z���܂���?", vbYesNo + vbQuestion, "���������쐬:" & ActiveSheet.Name)
        End If
        
        If overWrite = vbNo Then
            Exit Sub
        Else
            Set targetCell = Cells(lastRow, creditAmountColumn - 1)
        End If
    Else
        Set targetCell = Cells(lastRow + 1, creditAmountColumn - 1)
    End If
    
    ' // ���v���z�̌v�Z
    Dim totalAmount As Long
    totalAmount = WorksheetFunction.Sum(Range(Cells(firstRow, creditAmountColumn), Cells(targetCell.Row - 1, creditAmountColumn)))
    
    '// �ؕ��̖�����p(����)�̌v�Z
    Cells(firstRow, debitAmountColumn).value = totalAmount - _
        WorksheetFunction.Sum(Range(Cells(firstRow + 1, debitAmountColumn), Cells(targetCell.Row - 1, debitAmountColumn)))
    
    '// �ؕ��E�ݕ��̍��v�\��
    With targetCell
        .value = "�ݕ����v�z"
        .HorizontalAlignment = xlCenter
        .Offset(0, 1).value = totalAmount
        
        With .Offset(0, -2)
            .Formula = "�ؕ����v�z"
            .HorizontalAlignment = xlCenter
        End With
        .Offset(0, -1).value = totalAmount
    End With
    
    '// �`�F�b�N�{�b�N�X�쐬
    If overWrite <> vbYes And needCheckBox Then
        Dim chkController As New checkBoxController
        
        With targetCell
        '// chkContoller.add  [�ǉ�����Z��],[�`�F�b�N�{�b�N�X�̖��O]
            chkController.add Cells(.Row, .Offset(, 2).Column), "chk" & .Row
            Set chkController = Nothing
        End With
    End If
    
    Set targetCell = Nothing
    
    Range(Columns(debitAmountColumn - 1), Columns(creditAmountColumn)).EntireColumn.AutoFit

End Sub
'/**
' * ���͂��ꂽ�f�[�^���N���A
' * @params debitAmountColumn  �ؕ����z��
' * @params creditAmountColumn �ݕ����z��
' * @params headerRow          �w�b�_�[�s
' * @params firstClearDebitRow �ؕ��̃N���A���s���ŏ��̍s
' * @params startClearRow      �N���A���J�n����s
' * @params existCheckBox      �`�F�b�N�{�b�N�X�����݂��邩
'**/
Sub clearData(ByVal debitAmountColumn As Long, ByVal creditAmountColumn As Long, ByVal headerRow As Long, ByVal firstClearDebitRow As Long, ByVal startClearRow As Long, Optional ByVal existCheckBox As Boolean = True)
    

    Dim boundaryRow As Long
    boundaryRow = Application.InputBox("���s�ڂ̃f�[�^����N���A���邩���͂��Ă��������B", "���������쐬:" & ActiveSheet.Name, startClearRow, Type:=1)
    
    Application.ScreenUpdating = False
    
    If boundaryRow = 0 Then
        Exit Sub
    ElseIf boundaryRow <= 0 Then
        MsgBox "���͂��ꂽ�l�������ł��B", vbQuestion, "���������쐬:" & ActiveSheet.Name
        Exit Sub
    End If
    
    Dim lastRow As Long: lastRow = Cells(Rows.Count, creditAmountColumn - 1).End(xlUp).Row
    
    '// �ؕ��̃N���A
    Cells(headerRow + 1, debitAmountColumn).ClearContents
    Range(Cells(firstClearDebitRow, debitAmountColumn - 1), Cells(lastRow, debitAmountColumn)).ClearContents
        
    '// �`�F�b�N�{�b�N�X�̃`�F�b�N����
    Dim i As Long
    
    If existCheckBox Then
        For i = ActiveSheet.CheckBoxes.Count To 1 Step -1
            ActiveSheet.CheckBoxes(i).value = False
        Next
    End If
    
    '// �ݕ��N���A&�`�F�b�N�{�b�N�X�폜
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
'// �w��͈̔͂̈��
Private Sub printBill(ByVal targetRange As Range, ByVal msgTitle As String, Optional ByVal printOrientation As Long = xlPortrait)
    
    Application.ScreenUpdating = False
    
    '���b�Z�[�W��\�����Ċm�F
    If MsgBox("������܂��B��낵���ł���?", vbYesNo + vbQuestion, msgTitle) = vbNo Then
        Exit Sub
    End If
    
    '����ݒ�
    With ActiveSheet.PageSetup
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = 1
        .Orientation = printOrientation
    End With
    
    '���
    targetRange.PrintOut
    
End Sub
'/**
' * ���������t�@�C�����w��̃t�H���_�ɕۑ�
' * @params companyFolderPath   �ۑ���t�H���_(��Ж��܂�)
' * @params cutoffDate   ���ߓ�
' * @params paymentDate  �x������
' * @params ompanyName   ��Ж�
 '**/
Public Sub registerFile(ByVal companyFolderPath As String, ByVal cutoffDate As Date, ByVal paymentDate As Date, ByVal companyName As String)
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    Dim fso As New FileSystemObject
    
    '// companyFolderPath�Ŏw�肵���t�H���_��������Δ�����
    If fso.FolderExists(companyFolderPath) = False Then
        MsgBox "�ۑ���̃t�H���_�����݂��܂���B" & vbLf & "�u�ݒ�v�V�[�g�ŕۑ���t�H���_��ݒ肵�Ă��������B", vbQuestion, "���������쐬:" & companyName
        GoTo KILL
    End If
    
    '// ���ߓ�����ۑ���t�H���_���𔻒�
    '// ��:2022�N3��31�� �� 2021�N4���`3��
    '//    2022�N4��30�� �� 2022�N4���`3��
    Dim folderPath As String
    
    If Month(cutoffDate) >= 1 And Month(cutoffDate) <= 3 Then
        folderPath = companyFolderPath & "\" & Year(cutoffDate) - 1 & "�N4���`3��"
    Else
        folderPath = companyFolderPath & "\" & Year(cutoffDate) & "�N4���`3��"
    End If
    
    '// �ۑ���t�H���_��������΍쐬
    If fso.FolderExists(folderPath) = False Then
        fso.CreateFolder folderPath
    End If
    
    '// �ۑ�����t�@�C����
    Dim fileName As String
    fileName = folderPath & "\" & companyName & " " & Year(cutoffDate) & "�N" & Month(cutoffDate) & "����(" & Month(paymentDate) & "���x��).xlsm"
    
    '// ���Ƀt�@�C��������ꍇ�͏㏑�����邩�m�F
    If fso.FileExists(fileName) Then
        If MsgBox("���Ƀt�@�C��������܂����A�㏑�����܂���?", vbYesNo + vbQuestion, "���������쐬:" & companyName) = vbNo Then
            GoTo KILL
        End If
    End If
    
    '// ���Ƀt�@�C��������A�J���Ă���ƕۑ��ł��Ȃ����ߊJ���Ă��������
    Dim wb As Workbook
    For Each wb In Workbooks
        If wb.Name = fso.GetFileName(fileName) Then
            wb.Close
        End If
    Next
    
    ActiveWorkbook.SaveCopyAs fileName
    Application.Calculate
    
    MsgBox "�o�^���������܂����B", Title:="���������쐬:" & companyName

KILL:
    Set fso = Nothing

End Sub

'// �ۑ���t�H���_���ݒ�
Public Sub setFolderPath(companyName As String)

    With Application.FileDialog(msoFileDialogFolderPicker)
        .InitialFileName = "G:"
        .AllowMultiSelect = False
        .Title = "�ۑ���t�H���_�I��:" & companyName
        
        If .Show Then
            Sheets("�ݒ�").Cells(2, 3).value = .SelectedItems(1)
            
            MsgBox "�ۑ����ύX���܂����B", Title:="���������쐬:" & companyName
        End If
    End With
    
End Sub


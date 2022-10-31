Attribute VB_Name = "Main"
'// ���������쐬�̃��C�����W���[��
Option Explicit

'// �\���H�̂��߂̃t�H�[���N��
Public Sub openFormToProcessChart()

    Sheets("mode").Cells(1, 1).Value = "PROCESS_CHART"
    frmCompany.Show
    
End Sub

'// ���C�����[�`��(�\���H)
Public Sub processChart(ByVal company As String)

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    '// �r���ŃG���[�ɂȂ������ɓ\��t���Ȃ����ɂȂ�Ȃ��悤�ɕ\�̍ŏ��̏�Ԃ𑼂̃V�[�g�ɕێ�����
    Sheets("�Αӎx���T���ꗗ�\").Cells.Copy Sheets("tmp").Cells(1, 1)
    
    Dim companyRow As Long: companyRow = WorksheetFunction.Match(company, Sheets("�ݒ�").Columns(1), 0)
    Dim headerRow As Long: headerRow = Sheets("�Αӎx���T���ꗗ�\").Cells(1, 3).End(xlDown).Row

    Dim cc As New ChartController
    Dim vc As New ValueController

    '// �s�v�ȕ������폜
    If deleteUnnecessaryColumns(companyRow, headerRow, cc) = False Then
        Call resetChart
        GoTo Kill
    End If
    
    '// �u����v�ɂ��u�����v�ɂ��܂܂Ȃ���ړ�
    If moveIndependentColumns(companyRow, headerRow, cc) = False Then
        Call resetChart
        GoTo Kill
    End If
    
    '// ������쐬
    If createOfficeColumn(companyRow, headerRow, cc, vc) = False Then
        Call resetChart
        GoTo Kill
    End If
    
    '// �����쐬
    If createFieldWorkColumn(companyRow, headerRow, cc) = False Then
        Call resetChart
        GoTo Kill
    End If
    '// �x�����z�s�쐬
    Call createBasicSalaryRow(headerRow)
    
    '// �Αӎx���T���ꗗ�\�̌����ڒ���
    With Sheets("�Αӎx���T���ꗗ�\")
        Dim mainLastColumn As Long
        
        If Sheets("�ݒ�").Cells(companyRow, 3).Value = "" And Sheets("�ݒ�").Cells(companyRow, 4).Value = "" Then
            mainLastColumn = 0
        ElseIf Sheets("�ݒ�").Cells(companyRow, 3).Value <> "" And Sheets("�ݒ�").Cells(companyRow, 4).Value = "" Then
            mainLastColumn = WorksheetFunction.Match("����", .Rows(headerRow), 0)
        Else
            mainLastColumn = cc.searchDepartmentColumn(Int(Split(Sheets("�ݒ�").Cells(2, 4).Value, "-")(0)), 5)
        End If
        
        If mainLastColumn <> 0 Then
            .Range(Cells(headerRow, 1), .Cells(.Cells(Rows.Count, 1).End(xlUp).Row, mainLastColumn)).BorderAround Weight:=xlMedium
        End If
        
        .Range(.Rows(5), .Rows(6)).HorizontalAlignment = xlCenter
        
    End With
    
    '// �T���̕\�C��
    Call processDeductionChart(vc)
    
    Sheets("�Αӎx���T���ꗗ�\").Activate
    Cells(1, 1).Select
    
    MsgBox "�������������܂����B", Title:=ThisWorkbook.Name
    
Kill:
    Set cc = Nothing
    Set vc = Nothing
    
    Application.DisplayAlerts = True
    
End Sub

'// �s�v�ȗ���폜
Private Function deleteUnnecessaryColumns(ByVal companyRow As Long, ByVal headerRow As Long, cc As ChartController) As Boolean

    deleteUnnecessaryColumns = False

    If Sheets("�ݒ�").Cells(companyRow, 5).Value = "" Then
        deleteUnnecessaryColumns = True
        Exit Function
    End If

    Dim unnecessaryDepartments As Variant: unnecessaryDepartments = Split(Sheets("�ݒ�").Cells(companyRow, 5).Value, "-")
    Dim i As Long
    Dim targetColumn As Long
    
    For i = 0 To UBound(unnecessaryDepartments)
        targetColumn = cc.searchDepartmentColumn(unnecessaryDepartments(i), headerRow)
        If targetColumn = 0 Then: Exit Function
        
        Sheets("�Αӎx���T���ꗗ�\").Columns(targetColumn).Delete xlToLeft
    Next
    
    deleteUnnecessaryColumns = True
            
End Function

'// �u�����v�ɂ��u����v�ɂ��܂܂Ȃ�����ړ�
Private Function moveIndependentColumns(ByVal companyRow As Long, ByVal headerRow As Long, cc As ChartController) As Boolean

    moveIndependentColumns = False
    
    With Sheets("�ݒ�")
    
        If .Cells(companyRow, 4).Value = "" Then
            moveIndependentColumns = True
            Exit Function
        End If
    
        '// �u�����v�ɂ��u����v�ɂ��܂܂Ȃ���
        Dim independentColumns As Variant: independentColumns = Split(.Cells(companyRow, 4).Value, "-")
    End With
    
    Dim i As Long
    Dim targetColumn As Long
    
    With Sheets("�Αӎx���T���ꗗ�\")
        .Activate
    
        For i = 0 To UBound(independentColumns)
            
            If independentColumns(i) = 0 Then: GoTo Continue
            
            targetColumn = cc.searchDepartmentColumn(independentColumns(i), headerRow)
            If targetColumn = 0 Then: Exit Function
            
            '// ��ړ�
            cc.moveColumn targetColumn, 3
Continue:
        Next
    End With
    
    moveIndependentColumns = True
            
End Function

'// ������쐬
Private Function createOfficeColumn(ByVal companyRow As Long, ByVal headerRow As Long, cc As ChartController, vc As ValueController) As Boolean

    createOfficeColumn = False
    
    '// ������𔻕ʂ��邽�߂̃����_���̃R�[�h�����Z�b�g
    Sheets("mode").Cells(1, 2).Value = ""
    
    With Sheets("�ݒ�")
    '// ������Ɋ܂ޕ�����������Δ�����
        If .Cells(companyRow, 3).Value = "" Then
            createOfficeColumn = True
            Exit Function
        End If
        
        '// �u�����v�Ɋ܂ޕ����R�[�h
        Dim officeDepartments As Variant: officeDepartments = Split(.Cells(companyRow, 3).Value, "-")
    End With
            
    With Sheets("�Αӎx���T���ꗗ�\")
        .Columns(3).Insert xlToRight
        Dim officeCode As Long: officeCode = vc.generate8DigitsNumber
        .Cells(headerRow, 3).Value = officeCode & " ����"
        Sheets("mode").Cells(1, 2).Value = officeCode
        
        '//�u�����v�Ɋ܂ޕ����R�[�h�������͂��ꂽ��ԍ����i�[�����z����쐬
        Dim targetColumns As Variant: targetColumns = departmentCodes2ColumnNumbers(officeDepartments, headerRow, cc)
        If targetColumns(0) = "false" Then: Exit Function
        
        '// �l���v�Z
        Dim totalAmount As Long: totalAmount = sumPeople(targetColumns, headerRow, cc)
        If totalAmount = -1 Then: Exit Function
        .Cells(headerRow + 1, 3).Value = "�y �v " & totalAmount & "�� �z"
        
        '// �ʋΎ蓖�����v
        Dim sumFormula As String: sumFormula = sumAmountOfMoney(targetColumns, headerRow, headerRow + 2, cc)
        If sumFormula = "false" Then: Exit Function
        .Cells(headerRow + 2, 3).Formula = sumFormula
        
        '// ���̋��z�����v���o��悤��Autofill
        .Cells(headerRow + 2, 3).AutoFill .Range(.Cells(headerRow + 2, 3), .Cells(.Cells(Rows.Count, 1).End(xlUp).Row, 3))
        .Columns(3).Copy
        .Columns(3).PasteSpecial xlPasteValues
        
        '// �u�����v�Ɋ܂ޕ����̃w�b�_�[�̐F�ύX
        Dim i As Long
        For i = 0 To UBound(targetColumns)
            Cells(headerRow, Int(targetColumns(i))).Font.Color = RGB(51, 153, 51)
        Next
        
    End With

    createOfficeColumn = True

End Function

'// �����쐬
Private Function createFieldWorkColumn(ByVal companyRow As Long, ByVal headerRow As Long, cc As ChartController) As Boolean

    createFieldWorkColumn = False

    If Sheets("�ݒ�").Cells(companyRow, 6).Value = False Then
        createFieldWorkColumn = True
        Exit Function
    End If
    
    With Sheets("�Αӎx���T���ꗗ�\")
        .Columns(3).Insert xlToRight
        .Cells(headerRow, 3).Value = "����"
        
        '/**
         '* �l���v�Z
        '**/
        
        '// ����ȊO�̕����R�[�h(�u�����v�ɂ��u����v�ɂ��܂߂Ȃ������Ǝ���
        Dim officeCode As String: officeCode = ""
        If Sheets("mode").Cells(1, 2).Value <> "" Then
            officeCode = "-" & Sheets("mode").Cells(1, 2).Value
        End If
        
        Dim notFieldWorkDepartments As Variant: notFieldWorkDepartments = Split(Sheets("�ݒ�").Cells(companyRow, 4).Value & officeCode, "-")
        
        '// ����ȊO�̕����R�[�h���R�[�h�����͂���Ă����ԍ����i�[�����z��ɕϊ�
        Dim notFieldWorkColumns As Variant: notFieldWorkColumns = departmentCodes2ColumnNumbers(notFieldWorkDepartments, headerRow, cc)
        If notFieldWorkColumns(0) = "false" Then: Exit Function
        
        '//���v���猻��ȊO�̐l���������Č���̐l�������߂�
        Dim vc As New ValueController
        Dim numberOfFieldWorkers As Long
        
        numberOfFieldWorkers = cc.countPeople(2, headerRow, vc) - sumPeople(notFieldWorkColumns, headerRow, cc)
        .Cells(headerRow + 1, 3).Value = "�y �v " & numberOfFieldWorkers & "�� �z"
        
        Set vc = Nothing
        
        '/**
         '* �e����z����
        '**/
          
        '// �ʋΎ蓖�v�Z
        Dim fieldWorkFormula As String: fieldWorkFormula = "=B" & headerRow + 2
        Dim i As Long
        Dim columnAlphabet As String
        
        For i = 0 To UBound(notFieldWorkColumns)
            columnAlphabet = vc.columnNumber2Alphabet(Int(notFieldWorkColumns(i)))
            fieldWorkFormula = fieldWorkFormula & "-" & columnAlphabet & headerRow + 2
        Next
        
        .Cells(headerRow + 2, 3).Formula = fieldWorkFormula
        
        '// �e����z��AutoFill�Ŕ��f�����A���l��
        .Cells(headerRow + 2, 3).AutoFill .Range(.Cells(headerRow + 2, 3), .Cells(.Cells(Rows.Count, 1).End(xlUp).Row, 3))
        .Columns(3).Copy
        .Columns(3).PasteSpecial xlPasteValues
        
        '// ������̃w�b�_�[�̃R�[�h�폜
        If Sheets("�ݒ�").Cells(companyRow, 3).Value <> "" Then
            .Cells(headerRow, Int(notFieldWorkColumns(UBound(notFieldWorkColumns)))).Value = "����"
        End If
        
    End With
    
    createFieldWorkColumn = True
    
End Function

'/**
 '* �����̕����R�[�h���i�[���ꂽ�z�񂩂畔���R�[�h�����͂��ꂽ��ԍ����i�[�����z���Ԃ�
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
 '* ���x������ʋΎ蓖�����������z��\������u�x�����z�v�s���쐬
'**/
Private Sub createBasicSalaryRow(ByVal headerRow As Long)

    With Sheets("�Αӎx���T���ꗗ�\")
        .Rows(headerRow + 3).Insert xlDown
        .Cells(headerRow + 3, 1).Value = "�x�����z"
        
        .Cells(headerRow + 3, 2).Formula = "=B" & headerRow + 4 & "-B" & headerRow + 2
        .Cells(headerRow + 3, 2).AutoFill .Range(.Cells(headerRow + 3, 2), Cells(headerRow + 3, .Cells(headerRow, Columns.Count).End(xlToLeft).Column))
    
        .Range(.Cells(headerRow + 4, 1), Cells(headerRow + 4, .Cells(headerRow, Columns.Count).End(xlToLeft).Column)).Borders(xlEdgeBottom).LineStyle = xlDouble
    
    End With

End Sub

'/**
' * �w��̕����̓����s�̋��z�𑫂��Z����G�N�Z���̎���Ԃ�
' * @params targetDepartments ���z�����������R�[�h
'**/
Private Function sumAmountOfMoney(targetColumns As Variant, headerRow As Long, ByVal targetRow As Long, cc As ChartController) As String

    Dim i As Long
    Dim vc As New ValueController
    
    Dim returnFormula As String: returnFormula = "=SUM("
    
    '// ��ԍ����A���t�@�x�b�g�ɕϊ���������
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

'// ���������̍��v�l�������߂�
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

'// �u�T���v�̃V�[�g�̕\���H
Private Sub processDeductionChart(vc As ValueController)

    Dim sumRow As Long
    
    With Sheets("�T��")
        .Activate
    
        '// ���v���z��0�̍��ڂ��폜
        On Error Resume Next
        sumRow = WorksheetFunction.Match("�y ���v*", .Columns(1), 0)
        On Error GoTo 0
        
        If sumRow = 0 Then: Exit Sub
        
        Dim headerRow As Long: headerRow = Sheets("�T��").Cells(1, 3).End(xlDown).Row
        
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
        
        '// �S���ڂ̍��v��0�~�̐l���폜
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

'// �r���ŃG���[�ɂȂ������ɕ\���ŏ��̏�Ԃɖ߂�
Private Sub resetChart()

    Sheets("tmp").Cells.Copy Sheets("�Αӎx���T���ꗗ�\").Cells(1, 1)

End Sub

'// �\�\��t���̂��߂̃t�H�[���N��
Public Sub openFormToPasteChart()

    Sheets("mode").Cells(1, 1).Value = "PASTE_CHART"
    frmCompany.Show
    
End Sub

'// �\��t��
Public Sub pasteChart(ByVal company As String)

    Application.ScreenUpdating = False

    Dim path As String
    
    With Sheets("�ݒ�")
        path = .Cells(WorksheetFunction.Match(company, .Columns(1), 0), 2).Value
    End With
    
    If path = "" Then
        MsgBox "�\��t���悪�ݒ肳��Ă��܂���B", Title:=ThisWorkbook.Name
    End If
    
    Dim fso As New FileSystemObject
    
    If fso.FileExists(path) = False Then
        MsgBox "�\��t����t�@�C����������܂���ł����B" & vbLf & "�ݒ��ύX���Ă��������B" & vbLf & vbLf & "�y���ݐݒ蒆�̓\��t����z" & vbLf & path, vbQuestion, ThisWorkbook.Name
        GoTo Kill
    End If
    
    Workbooks.Open path
    
    Dim fileName As Variant: fileName = Split(path, "\")
    fileName = fileName(UBound(fileName))
    
    ThisWorkbook.Sheets("�Αӎx���T���ꗗ�\").Cells.Copy Workbooks(fileName).Sheets("�Αӎx���T���ꗗ�\").Cells(1, 1)
    ThisWorkbook.Sheets("�T��").Cells.Copy Workbooks(fileName).Sheets("�T��").Cells(1, 1)
    
    MsgBox "�\��t�����������܂����B", Title:=ThisWorkbook.Name
    
    ThisWorkbook.Close True

Kill:
    Set fso = Nothing
    
End Sub

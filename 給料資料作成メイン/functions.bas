Attribute VB_Name = "functions"
'// �V�[�g�Ŏg�p����֐����`���郂�W���[��
Option Explicit

'// �V�[�g�Ŏg�p����֐����֐��_�C�A���O�{�b�N�X�ɓo�^
Public Sub registerFunctions()

    Application.MacroOptions _
        Macro:="PAYMENT", Description:="����(��:�Љ�ی���)�ƑΏە���������z�����߂�֐��ł��B", _
        Category:="���������쐬", ArgumentDescriptions:=Array("����(��:�Љ�ی���)", "�Ώە����R�[�h�܂��͕�����", "�V�[�g�����Αӎx���T���ꗗ�\�̏ꍇ�͏ȗ��\�ł�� ")

End Sub

'/**
' * �T�����z���u�Αӎx���T���ꗗ�\�v����v�Z
' * @params targetType �T������(���N�ی����Ȃ�)
' * @params targetCode �T���Ώە����̃R�[�h�������͖��O

Public Function PAYMENT(ByVal targetType As Variant, targetCode As Variant, Optional sheetName As String = "�Αӎx���T���ꗗ�\") As Long
Attribute PAYMENT.VB_Description = "����(��:�Љ�ی���)�ƑΏە���������z�����߂�֐��ł��B"
Attribute PAYMENT.VB_ProcData.VB_Invoke_Func = " \n20"

    Application.Volatile

   On Error Resume Next
    
    '// ���ڂ̍s�ԍ�
    Dim targetRow As Long
    
    If sheetName = "�Αӎx���T���ꗗ�\" Then
        targetRow = WorksheetFunction.Match("*" & targetType & "*", Sheets(sheetName).Columns(1), 0)
    ElseIf sheetName = "�T��" Then
        targetRow = WorksheetFunction.Match("*" & targetCode & "*", Sheets(sheetName).Columns(2), 0)
    End If
    
    '// �Ώە����̗�ԍ�
    Dim targetColumn As Long
    
    If sheetName = "�Αӎx���T���ꗗ�\" Then
        targetColumn = WorksheetFunction.Match("*" & targetCode & "*", Sheets(sheetName).Rows(5), 0)
    ElseIf sheetName = "�T��" Then
        targetColumn = WorksheetFunction.Match("*" & targetType & "*", Sheets(sheetName).Rows(5), 0)
    End If
    
    On Error GoTo 0
    
    '// �T�����ځE�T���Ώە�����������Ȃ��A�������͂��̃Z���̒l�������o�Ȃ��ꍇ�͔�����
    If targetRow = 0 Or targetColumn = 0 Then
        PAYMENT = 0
        Exit Function
    ElseIf IsNumeric(Sheets(sheetName).Cells(targetRow, targetColumn).value) = False Then
        PAYMENT = 0
        Exit Function
    End If
        
    PAYMENT = Sheets(sheetName).Cells(targetRow, targetColumn).value

End Function
'// ��s����o�͂��ꂽ�f�[�^�̓��t����x�������v�Z
Public Function PAYDAY(ByVal dateOfBugyo As Variant) As Date
    
    Application.Volatile
    
    '1 ��s�̓��t��"�N" �ł킯�A�N�ƌ������߂�
    dateOfBugyo = Split(dateOfBugyo, "�N")
    
    '// �a��𐼗�ɕϊ�
    Dim yearOfBugyo As Long: yearOfBugyo = Format(dateOfBugyo(0) & "�N1��1��", "yyyy")
    '// �����擾
    Dim monthOfbugyo As Long: monthOfbugyo = Val(Split(dateOfBugyo(1), "��")(0))
    dateOfBugyo = DateSerial(yearOfBugyo, monthOfbugyo, 20)
    
    '2 �y���j���Əd�Ȃ�������t��1���O�ɂ��āA�����ɂȂ�܂ŌJ��ւ���
    Dim result As Boolean, i As Long
    i = 1
    Do Until result = True
        If Weekday(dateOfBugyo) = 1 Or Weekday(dateOfBugyo) = 7 Then
            dateOfBugyo = DateSerial(yearOfBugyo, monthOfbugyo, 20 - i)
            result = False
            i = i + 1
        ElseIf Application.WorksheetFunction.CountIf(Sheets("�ݒ�").Range("A:A"), dateOfBugyo) >= 1 Then
            dateOfBugyo = DateSerial(yearOfBugyo, monthOfbugyo, 20 - i)
            result = False
            i = i + 1
        Else
            result = True
        End If
    Loop
    PAYDAY = dateOfBugyo

End Function

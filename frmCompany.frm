VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmCompany 
   Caption         =   "��БI��"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "frmCompany.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "frmCompany"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'// ��Ж���I������t�H�[��
Option Explicit

'// �t�H�[���N����
Private Sub UserForm_Initialize()

    Dim i As Long

    With cmbCompany
        .Style = fmStyleDropDownCombo
    
        For i = 2 To Sheets("�ݒ�").Cells(Rows.Count, 1).End(xlUp).Row
            .AddItem Sheets("�ݒ�").Cells(i, 1).Value
            
            If InStr(1, Sheets("�Αӎx���T���ꗗ�\").Cells(2, 2).Value, Sheets("�ݒ�").Cells(i, 1).Value) > 0 Then
                .Value = Sheets("�ݒ�").Cells(i, 1).Value
            End If
        Next
        
        .Style = fmStyleDropDownList
    End With
    
End Sub

'// �u���s�v���������Ƃ�
Private Sub cmdEnter_Click()

    If cmbCompany.Value = "" Then
        MsgBox "��Ж���I�����Ă��������B", vbQuestion, ThisWorkbook.Name
        Exit Sub
    End If

    Select Case Sheets("mode").Cells(1, 1).Value
        Case "SET_PATH"
            Call setPath(cmbCompany.Value)
            Exit Sub
        Case "PROCESS_CHART"
            Call processChart(cmbCompany.Value)
        Case "PASTE_CHART"
            Call pasteChart(cmbCompany.Value)
    End Select
        
    Unload Me
    
End Sub

'// �u����v���������Ƃ��̏���
Private Sub cmdClose_Click()

    Unload Me

End Sub

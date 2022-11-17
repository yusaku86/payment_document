Attribute VB_Name = "setting"
Option Explicit

'// �\��t����̃t�@�C����ݒ�
Public Sub setPath(ByVal company As String)

    Dim path As String
    
    With Application.FileDialog(msoFileDialogFilePicker)
        .AllowMultiSelect = False
        .Title = "�\��t����t�@�C���̐ݒ�(" & company & ")"
        .InitialFileName = "G:\"
                
        If .Show Then
            path = .SelectedItems(1)
        Else
            Exit Sub
        End If
    End With
    
    With Sheets("�ݒ�")
        .Cells(WorksheetFunction.Match(company, .Columns(1), 0), 2).Value = path
        MsgBox "�\��t����t�@�C����ύX���܂����B", Title:=ThisWorkbook.Name
    End With
    
End Sub

'// ���[�U�[�t�H�[���N��(�t�@�C���p�X�ݒ�)
Public Sub openFormToSetPath()

    Sheets("mode").Cells(1, 1).Value = "SET_PATH"
    frmCompany.Show
    
End Sub

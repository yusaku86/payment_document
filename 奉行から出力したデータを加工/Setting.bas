Attribute VB_Name = "setting"
Option Explicit

'// 貼り付け先のファイルを設定
Public Sub setPath(ByVal company As String)

    Dim path As String
    
    With Application.FileDialog(msoFileDialogFilePicker)
        .AllowMultiSelect = False
        .Title = "貼り付け先ファイルの設定(" & company & ")"
        .InitialFileName = "G:\"
                
        If .Show Then
            path = .SelectedItems(1)
        Else
            Exit Sub
        End If
    End With
    
    With Sheets("設定")
        .Cells(WorksheetFunction.Match(company, .Columns(1), 0), 2).Value = path
        MsgBox "貼り付け先ファイルを変更しました。", Title:=ThisWorkbook.Name
    End With
    
End Sub

'// ユーザーフォーム起動(ファイルパス設定)
Public Sub openFormToSetPath()

    Sheets("mode").Cells(1, 1).Value = "SET_PATH"
    frmCompany.Show
    
End Sub

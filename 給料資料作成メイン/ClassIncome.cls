VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClassIncome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'指定の範囲にチェックボックスを追加
 Sub add(targetRange As Range, ByVal chkName As String)
    
    'chkNameと同じ名前のチェックボックスが既にあるか確認し、あればチェックボックスを追加しない
    Dim i As Long
    With ActiveSheet
        For i = 1 To .Shapes.Count
            If .Shapes(i).name = chkName Then
                Exit Sub
            End If
        Next
    End With
    
    'チェックボックス追加
    With targetRange
        With ActiveSheet.CheckBoxes.add(.Left, .Top, .Width, .Height)
            .Caption = ""
            .name = chkName
        End With
    End With

End Sub

'作成したブックを所定のフォルダに格納
Function Register(ByVal folderPath As String, ByVal filePath As String) As Boolean
    
    '0 メッセージにて確認
    Dim ans As VbMsgBoxResult: ans = MsgBox("登録してよろしいですか?", vbYesNo + vbQuestion)
    If ans = vbNo Then
        Register = False
        Exit Function
    End If
    
    '1 指定のフォルダが無ければ作成、あればそこに保存
    Dim myFSO As New FileSystemObject
    
    If myFSO.FolderExists(folderPath) = False Then
        myFSO.CreateFolder folderPath
    End If
    
    Set myFSO = Nothing
    
    On Error GoTo errorTrap
    
    ThisWorkbook.SaveAs filePath
    Register = True
    
    Exit Function
      
errorTrap:
    MsgBox "処理を中断しました。" & vbLf & "登録は完了していません。", vbExclamation
    Register = False
    Exit Function
    
End Function

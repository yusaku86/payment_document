VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "checkBoxController"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'// チェックボックスを制御するクラス
Option Explicit
'指定の範囲にチェックボックスを追加
Public Sub add(targetRange As Range, ByVal chkName As String)
    
    'chkNameと同じ名前のチェックボックスが既にあるか確認し、あればチェックボックスを追加しない
    If isExistChk(ActiveSheet, chkName) = True Then: Exit Sub
    
    'チェックボックス追加
    With targetRange
        With ActiveSheet.CheckBoxes.add(.Left, .Top, .Width, .Height)
            .Caption = ""
            .Name = chkName
        End With
    End With

End Sub

'// 指定の名前のチェックボックスが存在するか
Public Function isExistChk(ByVal targetSheet As Worksheet, ByVal chkName As String) As Boolean

    Dim tmpShape As Shape
    
    For Each tmpShape In targetSheet.Shapes
        If tmpShape.Name = chkName Then
            isExistChk = True
            Exit Function
        End If
    Next
            
    isExistChk = False
    
End Function

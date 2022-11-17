VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FormComboBoxController"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'// フォームのコンボボックスを制御するクラス
Option Explicit

'// 入力された値がリストから選択されたか(リストと完全一致するか)どうかを判断
Public Function isSelectedFromList(ByVal value As Variant, ByRef lists As Variant) As Boolean

    Dim i As Long
    
    For i = 0 To UBound(lists)
        If value = lists(i) Then
            isSelectedFromList = True
            Exit Function
        End If
    Next
    
    isSelectedFromList = False
    
End Function


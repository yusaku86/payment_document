VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Validator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'// バリデーションを行うクラス
Option Explicit

'// 空白かどうかを判定
Public Function required(ByVal value As Variant) As Boolean

    If Replace(Replace(value, " ", ""), "　", "") = "" Then
        required = False
    Else
        required = True
    End If

End Function

'// 依存関係の確認(ある値が入力されている場合に、他の特定の値が入力されているかを確認する)
Public Function requiredWith(ByRef dependencies As Variant, ByVal value As Variant) As Boolean
            
    Dim i As Long
    
    For i = 0 To UBound(dependencies)
        If Me.required(dependencies(i)) = True And Me.required(value) = False Then
            requiredWith = False
            Exit Function
        End If
    Next
    
    requiredWith = True

End Function

'// 値が指定の正規表現パターンにマッチするか
Public Function pregMatch(ByVal value As Variant, ByVal pregPattern As String, reg As RegExp) As Boolean

    With reg
        .Global = True
        .Pattern = pregPattern
        
        If .Test(value) Then
            pregMatch = True
        Else
            pregMatch = False
        End If
    End With
    
End Function




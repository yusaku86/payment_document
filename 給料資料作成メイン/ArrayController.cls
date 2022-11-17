VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ArrayController"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'// 配列を制御するクラス
Option Explicit

'// 配列の末尾に要素を追加
Public Sub arrayPush(ByRef targetArray As Variant, ByVal value As Variant)

    ReDim Preserve targetArray(UBound(targetArray) + 1)
    targetArray(UBound(targetArray)) = value

End Sub

'/**
 '* 配列から指定の要素を削除(インデックスで指定)
 '* 配列から指定の要素を削除した配列を返す
'**/
Public Function arrayRemoveIndex(ByRef targetArray As Variant, ByVal index As Long) As Variant

    Dim i As Long
    
    Dim returnArray() As Variant
    ReDim returnArray(0)
    returnArray(0) = "NotAdded"
    
    For i = 0 To UBound(targetArray)
        '// 削除対象の要素は跳ばす
        If i = index Then
            GoTo Continue
        End If
        
        If returnArray(0) = "NotAdded" Then
            returnArray(0) = targetArray(i)
        Else
            Me.arrayPush returnArray, targetArray(i)
        End If
        
Continue:
    Next
    
    arrayRemoveIndex = returnArray
    
End Function


'// 配列から特定の値をもつ要素を削除
Public Function arrayRemoveValue(ByRef targetArray As Variant, ByVal value As Variant) As Variant

    Dim i As Long
    
    Dim returnArray() As Variant
    ReDim returnArray(0)
    
    For i = 0 To UBound(targetArray)
        If targetArray(i) <> value Then
            Me.arrayPush returnArray, targetArray(i)
        End If
    Next
    
    arrayRemoveValue = returnArray

End Function


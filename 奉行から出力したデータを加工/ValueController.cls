VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ValueController"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'// 文字列関連の操作を制御するクラス
Option Explicit

'// replace関数を一括で行う
Public Function replaceAtOnce(ByVal targetValue As Variant, delimeters() As Variant, substituteValue As Variant) As Variant

    Dim i As Long
    Dim replacedValue As Variant: replacedValue = targetValue
    
    For i = 0 To UBound(delimeters)
        If delimeters(i) = "" Then: GoTo Continue
        
        replacedValue = Replace(replacedValue, delimeters(i), substituteValue)
Continue:
    Next
    
    replaceAtOnce = replacedValue

End Function

'// 列番号をアルファベットに変換
Public Function columnNumber2Alphabet(columnNumber As Long) As Variant

    columnNumber2Alphabet = Split(Columns(columnNumber).Address, ":$")(1)

End Function

'// ランダムの8桁の数字生成
Public Function generate8DigitsNumber() As Long

    Dim randomNumber As String
    Dim i As Long
    
    For i = 1 To 8
        Randomize
        randomNumber = randomNumber & Int(10 * Rnd)
    Next
    
    generate8DigitsNumber = Int(randomNumber)

End Function


VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ChartController"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'// 表の操作を制御するクラス
Option Explicit

'// 列を切り取り別の列に挿入
Public Sub moveColumn(ByVal targetColumn As Long, destinationColumn As Long)

    Columns(targetColumn).Cut
    Columns(destinationColumn).Insert xlToRight
    
End Sub

'// 部署コードから列番号を求める
Public Function searchDepartmentColumn(ByVal departmentCode As Variant, headerRow As Long)

    Dim targetColumn As Long: targetColumn = 0

    On Error Resume Next
    
    targetColumn = WorksheetFunction.Match(departmentCode & " *", Sheets("勤怠支給控除一覧表").Rows(headerRow), 0)
    
    On Error GoTo 0
    
    If targetColumn = 0 Then
        MsgBox "部署コード:" & departmentCode & "が見つかりません。" & vbLf & "設定シートで設定を変更してください。", vbQuestion, ThisWorkbook.Name
    End If
    
    searchDepartmentColumn = targetColumn

End Function

'// 部署の人数を求める
Public Function countPeople(ByVal targetColumn As Long, headerRow As Long, vc As ValueController) As Long

    Dim delimeters(3) As Variant
    delimeters(0) = " "
    delimeters(1) = "　"
    delimeters(2) = "【計"
    delimeters(3) = "名】"

    With Sheets("勤怠支給控除一覧表")
        countPeople = Int(vc.replaceAtOnce(.Cells(headerRow + 1, targetColumn).Value, delimeters, ""))
    End With
End Function



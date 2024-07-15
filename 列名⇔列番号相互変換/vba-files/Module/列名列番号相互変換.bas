Attribute VB_Name = "列名列番号相互変換"
Option Explicit

' getLastIdxNo関数を使用するために必要
Sub test_getLastIdxNo()
    Dim lastIdxNo As Integer
    
    ' 引数の列番号の列の最下行番号を取得
    lastIdxNo = getLastIdxNo(Worksheets("Sheet1"), 1)
    
    Debug.Print lastIdxNo
End Sub

' 列番号→列名
Function columnIdx2Name(ByVal colNum As Long) As String
    columnIdx2Name = Split(Columns(colNum).Address, "$")(2)
End Function

' 列名→列番号
Function columnName2Idx(ByVal colName As String) As Long
    columnName2Idx = Columns(colName).Column
End Function

' 引数の列番号の列の最下行番号を取得
Function getLastIdxNo(targetSheet, colNum)
    Dim lastIdxNo As Integer
    
    ' 引数の列番号の列の最下行番号を取得
    lastIdxNo = targetSheet.Cells(Rows.Count, colNum).End(xlUp).Row

    getLastIdxNo = lastIdxNo
End Function

Attribute VB_Name = "�񖼗�ԍ����ݕϊ�"
Option Explicit

' getLastIdxNo�֐����g�p���邽�߂ɕK�v
Sub test_getLastIdxNo()
    Dim lastIdxNo As Integer
    
    ' �����̗�ԍ��̗�̍ŉ��s�ԍ����擾
    lastIdxNo = getLastIdxNo(Worksheets("Sheet1"), 1)
    
    Debug.Print lastIdxNo
End Sub

' ��ԍ�����
Function columnIdx2Name(ByVal colNum As Long) As String
    columnIdx2Name = Split(Columns(colNum).Address, "$")(2)
End Function

' �񖼁���ԍ�
Function columnName2Idx(ByVal colName As String) As Long
    columnName2Idx = Columns(colName).Column
End Function

' �����̗�ԍ��̗�̍ŉ��s�ԍ����擾
Function getLastIdxNo(targetSheet, colNum)
    Dim lastIdxNo As Integer
    
    ' �����̗�ԍ��̗�̍ŉ��s�ԍ����擾
    lastIdxNo = targetSheet.Cells(Rows.Count, colNum).End(xlUp).Row

    getLastIdxNo = lastIdxNo
End Function

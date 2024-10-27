Attribute VB_Name = "Module1"
Option Explicit

Const POWER As Integer = 3

Dim LOOP_COUNT As Long

' ���C������
Sub Hoge()    
    ' ��ʕ`����I�t�ɂ���
    Application.ScreenUpdating = False
    
    ' "copy"�V�[�g�����݂��Ă���΍폜
    ' "result"�V�[�g�����݂��Ă���΍폜
    On Error Resume Next
    Application.DisplayAlerts = False
    Sheets("copy").Delete
    Sheets("result").Delete
    Application.DisplayAlerts = True
    
    ' "result"�V�[�g���쐬
    Sheets.Add(After:=Sheets(Sheets.Count)).Name = "result"
    Dim sheet_result As Worksheet
    Set sheet_result = Sheets("result")
    sheet_result.Cells(1, 1).Value = "�������e"
    sheet_result.Cells(2, 1).Value = "�Z�����g�������[�v����"
    sheet_result.Cells(3, 1).Value = "1�����z����g�������[�v����"
    sheet_result.Cells(4, 1).Value = "2�����z����g�������[�v����"
    sheet_result.Cells(1, 2).Value = "�l�ݒ菈������"
    sheet_result.Cells(1, 3).Value = "�R�s�[��������"
    
    ' "copy"�V�[�g���쐬
    Sheets.Add(After:=Sheets(Sheets.Count)).Name = "copy"
    Dim copy_sheet As Worksheet
    Set copy_sheet = Sheets("copy")

    Dim value_sheet As Worksheet
    Set value_sheet = Sheets("value")
    
    ' 10,100,1000��̃��[�v�������s��
    Dim i as Integer
    For i = 1 To POWER
        LOOP_COUNT = 10 ^ i        

        ' �S�Z�����N���A
        value_sheet.Cells.ClearContents
        copy_sheet.Cells.ClearContents
    
        Call ForLoopCells(value_sheet, copy_sheet, sheet_result)
    
        Call ForLoopRows(value_sheet, copy_sheet, sheet_result)
    
        Call ForLoopArrays(value_sheet, copy_sheet, sheet_result)
    Next i

    ' ��ʕ`����I���ɂ���
    Application.ScreenUpdating = True

    MsgBox "�������������܂����B"

End Sub
    
' �Z�����g�������[�v����
Sub ForLoopCells(value_sheet as Worksheet, copy_sheet as Worksheet, sheet_result as Worksheet)
    Dim i As Long
    Dim j As Long
    Dim colName As String
    Dim startTime As Double
    Dim endTime As Double
    Dim elapsedTime As Double

    ' value�V�[�g�ɒl��ݒ�
    startTime = Timer
    For i = 1 To LOOP_COUNT
        For j = 1 To LOOP_COUNT
            value_sheet.Cells(i, j).Value = i * j
        Next j
    Next i
    endTime = Timer
    elapsedTime = endTime - startTime
    sheet_result.Cells(2, 2).Value = elapsedTime

    ' copy�V�[�g�ɒl���R�s�[(�Z�����g�������[�v����)
    startTime = Timer
    For i = 1 To LOOP_COUNT
        For j = 1 To LOOP_COUNT
            copy_sheet.Cells(i, j).Value = value_sheet.Cells(i, j).Value
        Next j
    Next i
    endTime = Timer
    elapsedTime = endTime - startTime
    sheet_result.Cells(2, 3).Value = elapsedTime
End Sub

' 1�����z����g�������[�v����
Sub ForLoopRows(value_sheet as Worksheet, copy_sheet as Worksheet, sheet_result as Worksheet)
    Dim i As Long
    Dim j As Long
    Dim colName As String
    Dim startTime As Double
    Dim endTime As Double
    Dim elapsedTime As Double

    ' 1�����z����쐬
    Dim arr1 As Variant

    ' j�̒l(��ԍ�)��񖼂ɕϊ�
    colName = ColNumToColName(LOOP_COUNT, value_sheet)
    
    ' value�V�[�g�ɒl��ݒ�
    startTime = Timer
    For i = 1 To LOOP_COUNT
        ' �z��̏�����
        ReDim arr1(LOOP_COUNT - 1)
        For j = 1 To LOOP_COUNT
            arr1(j - 1) = i * j
        Next j

        ' A1���N�_�ɏo��
        value_sheet.Range("A" & i & ":" & colName & i).Value = arr1
    Next i
    endTime = Timer
    elapsedTime = endTime - startTime
    sheet_result.Cells(3, 2).Value = elapsedTime

    ' copy�V�[�g�ɒl���R�s�[(1�����z����g�������[�v����)
    startTime = Timer
    For i = 1 To LOOP_COUNT
        arr1 = value_sheet.Range("A" & i & ":" & colName & i).Value
        copy_sheet.Range("A" & i & ":" & colName & i).Value = arr1
    Next i
    endTime = Timer
    elapsedTime = endTime - startTime
    sheet_result.Cells(3, 3).Value = elapsedTime
End Sub

' 2�����z����g�������[�v����
Sub ForLoopArrays(value_sheet as Worksheet, copy_sheet as Worksheet, sheet_result as Worksheet)
    Dim i As Long
    Dim j As Long
    Dim colName As String
    Dim startTime As Double
    Dim endTime As Double
    Dim elapsedTime As Double

    ' 2�����z����쐬
    Dim arr2 As Variant

    ' j�̒l(��ԍ�)��񖼂ɕϊ�
    colName = ColNumToColName(LOOP_COUNT, value_sheet)
    
    ' �z��̏�����
    ReDim arr2(LOOP_COUNT - 1, LOOP_COUNT - 1)
    
    ' value�V�[�g�ɒl��ݒ�
    startTime = Timer
    For i = 1 To LOOP_COUNT
        For j = 1 To LOOP_COUNT
            arr2(i - 1, j - 1) = i * j
        Next j
    Next i    
    ' arr2��A1���N�_�ɏo��
    value_sheet.Range("A1:" & colName & LOOP_COUNT).Value = arr2
    endTime = Timer
    elapsedTime = endTime - startTime
    sheet_result.Cells(4, 2).Value = elapsedTime

    ' copy�V�[�g�ɒl���R�s�[
    startTime = Timer
    arr2 = value_sheet.Range("A1:" & colName & LOOP_COUNT).Value
    copy_sheet.Range("A1:" & colName & LOOP_COUNT).Value = arr2
    endTime = Timer
    elapsedTime = endTime - startTime
    sheet_result.Cells(4, 3).Value = elapsedTime
End Sub

' ��ԍ���񖼂ɕϊ�
Function ColNumToColName(colNum As Long, value_sheet as Worksheet) As String
    Dim colName As String
    colName = Split(value_sheet.Cells(1, colNum).Address, "$")(1)
    ColNumToColName = colName
End Function

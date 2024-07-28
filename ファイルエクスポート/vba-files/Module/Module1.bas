Attribute VB_Name = "Module1"
Option Explicit

Sub �{�^��1_Click()
    Dim targetPath As String
    
    ' �t�@�C���o�͐�
    targetPath = ActiveWorkbook.Worksheets(1).Cells(2, 3).Value
    
    Call export_file(ActiveWorkbook.Worksheets(2), targetPath)
    
    Debug.Print "End"

End Sub

' �V�[�g�̓��e��CSV�`��(UTF-8, LF)�ŏo�͂���
Sub export_file(targetWorksheet, targetPath)
    Dim maxRowNum As Long
    Dim targetRange As Range
    Dim i As Long
    Dim sheetName As String
    Dim fw As Variant
    Dim byteData() As Byte

    sheetName = targetWorksheet.Name

    ' �ŉ��s�̍s�ԍ����擾
    maxRowNum = getMaxRowNum(targetWorksheet)
    
    Set fw = CreateObject("ADODB.Stream")
    fw.Charset = "UTF-8"
    fw.Open
    
    ' �t�@�C���o�͑Ώ۔͈͂�����
    Set targetRange = targetWorksheet.Range("A1:B" & maxRowNum)
    
    ' A1B1A2B2...�̏��ő���
    For i = 1 To targetRange.Count
        
        ' B���������ɉ��s�R�[�h��t�^
        If i Mod 2 = 0 Then
            fw.WriteText targetRange.Item(i) & vbLf, 0
        Else
            fw.WriteText targetRange.Item(i) & ",", 0
        End If
        
    Next
    
    ' BOM�Ȃ�UTF8�쐬�̂��߂̍��
    fw.Position = 0
    ' adTypeBinary = 1
    fw.Type = 1
    fw.Position = 3

    byteData = fw.Read
    fw.Close

    fw.Open
    fw.Write byteData
    fw.SaveToFile targetPath & "\" & sheetName & ".csv", 2
    fw.Close
    
    Set fw = Nothing
    Set targetRange = Nothing

End Sub

' A��̍ŉ��s�ԍ����擾����
Function getMaxRowNum(targetWorksheet)
    Dim maxRowNum As Long
    
    maxRowNum = targetWorksheet.Cells(Rows.Count, 1).End(xlUp).Row

    getMaxRowNum = maxRowNum
End Function

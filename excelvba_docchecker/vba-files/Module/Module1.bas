Attribute VB_Name = "Module1"
Option Explicit

' �萔��`
' True�ɂ����Debug.Print���L��
Const IS_DEBUG As Boolean = True

' ���ʕϐ���`
' FSO�͋��ʒ�`�𗘗p
Dim FSO As Object

Sub docchecker()

    Dim rootDir As String
    Dim sheetName As String

    ' �I�u�W�F�N�g�̓v���V�[�W�����ł̐錾���K�v�Ȃ��߂����Ŏ��{
    Set FSO = CreateObject("Scripting.FileSystemObject")

    ' �t�H���_�I���_�C�A���O
    With Application.FileDialog(msoFileDialogFolderPicker)
        If .Show = True Then
            rootDir = .SelectedItems(1)
        End If
    End With

    Call DebugPrint(rootDir)

    ' �t�H���_�p�X�L�薳���`�F�b�N
    If rootDir = "" Then
        MsgBox "�N�_�ƂȂ�t�H���_��I�����Ă��������B"
        Exit Sub
    End If

    ' ���ʏo�͗p�V�[�g�̒ǉ�
    sheetName = addResultSheet()

    ' �t�H���_�z���̃t�@�C���𑖍�
    Call getFilesRecursive(rootDir, sheetName)
    
    ' �o�͌��ʃV�[�g����Ƀ`�F�b�N�����{
    Call checkDocument(sheetName)
    
    MsgBox ("�I��")

End Sub

' ���ʏo�͗p�V�[�g�̒ǉ�
Function addResultSheet()
    Dim NewWorkSheet As Worksheet
    Dim sheetName As String

    ' �V�[�g���͔N���������b
    sheetName = Format(Now, "yyyymmddhhnnss")

    Set NewWorkSheet = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    NewWorkSheet.Name = sheetName
    
    ' �w�b�_�s�̒ǉ�
    NewWorkSheet.Cells(2, columnName2Idx("B")).Value = "No"
    NewWorkSheet.Cells(2, columnName2Idx("C")).Value = "FilePath"
    NewWorkSheet.Cells(2, columnName2Idx("D")).Value = "Extention"
    NewWorkSheet.Cells(2, columnName2Idx("E")).Value = "FileName"
    NewWorkSheet.Cells(2, columnName2Idx("F")).Value = "���ŗ����V�[�g����"
    NewWorkSheet.Cells(2, columnName2Idx("G")).Value = "�Ő�"
    NewWorkSheet.Cells(2, columnName2Idx("H")).Value = "���œ�"
    NewWorkSheet.Cells(2, columnName2Idx("I")).Value = "�쐬�Җ�"
    NewWorkSheet.Cells(2, columnName2Idx("J")).Value = "�t�@�C�����ɔŐ�����"
    NewWorkSheet.Cells(2, columnName2Idx("K")).Value = "�t�@�C�����ɓ��t����"
    NewWorkSheet.Cells(2, columnName2Idx("L")).Value = "�V�t�@�C����"

    addResultSheet = sheetName
End Function

' �����ŗ^����ꂽ�p�X���ċA�I�ɑ�������
Sub getFilesRecursive(path, sheetName)
    Dim objFolder As Object
    Dim objFile As Object

    ' GetFolder(�t�H���_��).SubFolders�Ńt�H���_�z���̃t�H���_�ꗗ���擾
    For Each objFolder In FSO.GetFolder(path).SubFolders
        Call getFilesRecursive(objFolder.path, sheetName)
    Next

    ' GetFolder(�t�H���_��).Files�Ńt�H���_�z���̃t�@�C���ꗗ���擾
    ' ���ʂ��V�[�g�ɋL�^
    For Each objFile In FSO.GetFolder(path).Files
        Call writeFilePathToSheet(objFile, sheetName)
    Next

End Sub

' �������ʂ��V�[�g�ɓ]�L����
Sub writeFilePathToSheet(objFile, sheetName)
    Dim targetSheet As Worksheet
    
    Dim idx As Integer
    Dim idxNo As Integer

    Set targetSheet = ThisWorkbook.Worksheets(sheetName)
    
    ' idx = Excel�s�ԍ� , idxNo = �ʂ��ԍ�
    idx = getLastIdxNo(targetSheet, columnName2Idx("B")) + 1
    idxNo = idx - 2
    
    ' ���ʃV�[�g�ɂ͈ȉ���]�L����
    ' �t�@�C���p�X
    ' �g���q
    ' �t�@�C����
    targetSheet.Cells(idx, columnName2Idx("B")).Value = idxNo
    targetSheet.Cells(idx, columnName2Idx("C")).Value = objFile.path
    targetSheet.Cells(idx, columnName2Idx("D")).Value = getExtentionName(objFile.path)
    targetSheet.Cells(idx, columnName2Idx("E")).Value = getFileName(objFile.path)
End Sub

' �X�̃h�L�������g�ɂ��Ċm�F���s��
Sub checkDocument(sheetName)
    Dim targetSheet As Worksheet
    Dim lastRowNum As Integer
    Dim i As Integer
    Dim filePath As String
    Dim extention As String
    Dim fileName As String
    Dim tmpBook As Workbook
    Dim kaihanRirekiSheetNum As Integer
    Dim newFileName As String

    ' Word����p�̒�`
    Dim wordApp As Object
    Dim wordDoc As Object
    
    Set targetSheet = ThisWorkbook.Worksheets(sheetName)
    lastRowNum = getLastIdxNo(targetSheet, columnName2Idx("B"))
    Call DebugPrint(lastRowNum)
    
    ' �w�b�_�s��1�s������A�ŉ��s�܂ő���
    For i = 3 To lastRowNum
        filePath = targetSheet.Cells(i, columnName2Idx("C")).Value
        extention = targetSheet.Cells(i, columnName2Idx("D")).Value
        fileName = targetSheet.Cells(i, columnName2Idx("E")).Value
    
        Call DebugPrint(extention)

        ' �g���q�`�F�b�N
        ' Word�̏ꍇ
        If extention Like "doc*" Then
            'Word�����̏ꍇ�͓s�xWord�̋N����~���s��
            Set wordApp = CreateObject("Word.Application")
            wordApp.Visible = True
  
            Set wordDoc = wordApp.Documents.Open(fileName:=filePath, ReadOnly:=True)

            Call parseWordDoc(targetSheet, i, wordDoc)
            
            wordDoc.Close (False)
            wordApp.Quit

        ' Excel�̏ꍇ
        ElseIf extention Like "xls*" Then
            Call DebugPrint("xls")
        
            Set tmpBook = Workbooks.Open(fileName:=filePath, ReadOnly:=True)
            
            kaihanRirekiSheetNum = hasKaihanRirekiSheet(tmpBook)
            
            ' ���ŗ����V�[�g����
            If kaihanRirekiSheetNum > 0 Then
                targetSheet.Cells(i, columnName2Idx("F")).Value = "����"
                
                ' ���ŗ����V�[�g����Ő��ƋL���҂��擾���Č��ʂ�]�L
                Call parseKaihanRirekiSheet(targetSheet, i, tmpBook.Worksheets(kaihanRirekiSheetNum))
            
            ' ���ŗ����V�[�g�Ȃ� -> �������Ȃ�
            Else
                targetSheet.Cells(i, columnName2Idx("F")).Value = "�Ȃ�"
            End If
            
            tmpBook.Close (False)
        Else
            Call DebugPrint("Other:" & extention)
        End If
 
        ' ���l�[��
        ' �V�t�@�C�����𐶐�
        newFileName = getNewFileName(fileName, targetSheet, i)
    
    Next
    
End Sub

' �t�@�C�����ɔŐ��Ɠ��t���܂܂�邩�`�F�b�N���A�V�t�@�C�����𐶐�
Function getNewFileName(fileName As String, targetSheet As Worksheet, rowIdx As Integer)
    Dim tmpArr As Variant
    Dim arrLen As Integer
    Dim fileNameWithoutExtention As String
    Dim rirekiNo As String
    Dim kaihanDateStr As String
    Dim i As Integer
    Dim newFileName As String
    
    ' �t�@�C��������g���q������
    fileNameWithoutExtention = Mid(fileName, 1, InStrRev(fileName, ".") - 1)
    Call DebugPrint(fileNameWithoutExtention)

    ' "_"�Ńt�@�C�����𕪊����Ĕz��Ɋi�[
    tmpArr = Split(fileNameWithoutExtention, "_")
    
    ' �z�񒷂��v�Z
    arrLen = UBound(tmpArr) - LBound(tmpArr) + 1
    
    ' ����̏����ł͂Ȃ��i������̔z�񒷂�4�����j�ꍇ�͌��̃t�@�C������Ԃ�
    If arrLen < 4 Then
        targetSheet.Cells(rowIdx, columnName2Idx("L")).Value = fileName
        getNewFileName = fileName
        Exit Function
    Else
        ' �Ōォ��2�ځF�Ő��A�Ō�F���œ��t�Ƃ��Ēl���擾
        rirekiNo = tmpArr(arrLen - 2)
        kaihanDateStr = tmpArr(arrLen - 1)
        
        Call DebugPrint(rirekiNo & "," & kaihanDateStr)
        
        '�@�Ő��A���œ��t�Ƃ��ɓ��t�̏ꍇ�͐V�t�@�C�����𐶐����ĕԋp
        If IsNumeric(rirekiNo) = True And IsNumeric(kaihanDateStr) = True Then
            targetSheet.Cells(rowIdx, columnName2Idx("J")).Value = Split(rirekiNo, "")
            targetSheet.Cells(rowIdx, columnName2Idx("K")).Value = Split(kaihanDateStr, "")
            
            ' �Ő��A���œ��͉��ŗ����̒l����ɐ����B���ŗ����ɒl���Ȃ���Ό��̒l�Ő���
            If targetSheet.Cells(rowIdx, columnName2Idx("G")).Value = "" Or targetSheet.Cells(rowIdx, columnName2Idx("H")).Value = "" Then
                newFileName = fileName
                targetSheet.Cells(rowIdx, columnName2Idx("L")).Value = newFileName
                getNewFileName = newFileName
                Exit Function
            Else
                ' �Ő��A���œ��ȊO�͌��̒l�Ő���
                For i = 0 To arrLen - 3
                    newFileName = newFileName & tmpArr(i) & "_"
                Next
            
                ' �Ő�
                newFileName = newFileName & targetSheet.Cells(rowIdx, columnName2Idx("G")).Value & "_"
                
                ' ���œ��t
                newFileName = newFileName & Format(targetSheet.Cells(rowIdx, columnName2Idx("H")).Value, "yyyymmdd") & "_"
                
                ' �g���q
                newFileName = newFileName & "." & targetSheet.Cells(rowIdx, columnName2Idx("D")).Value
            
                targetSheet.Cells(rowIdx, columnName2Idx("L")).Value = newFileName
            
                getNewFileName = newFileName
                Exit Function
            End If
        End If
    End If
End Function

' Word����������Ő��Ɖ��ŎҖ����擾����
Sub parseWordDoc(targetSheet, i, wordDoc)
    Dim rirekiNo As String
    Dim kaihanDataStr As String
    Dim kaihanShaName As String
    Dim tbl As Object
    Dim j As Integer
    Dim k As Integer
    Dim tmpVal As String
    
    ' �������̕\�̐����`�F�b�N����B0�Ȃ牽�����Ȃ�
    If wordDoc.Tables.Count < 1 Then
        rirekiNo = "�\�Ȃ�"
        kaihanShaName = "�\�Ȃ�"
    Else
        ' �\������ꍇ��1�ڂ̕\�𑖍�����
        Set tbl = wordDoc.Tables(1)
        
        For j = 1 To tbl.Rows.Count
            For k = 1 To tbl.Columns.Count
                tmpVal = Left(tbl.Cell(j, k), Len(tbl.Cell(j, k)) - 2)
            
                If tmpVal <> "" And k = 1 Then
                    rirekiNo = tmpVal
                ElseIf tmpVal <> "" And k = 2 Then
                    kaihanDataStr = tmpVal
                ElseIf tmpVal <> "" And k = 3 Then
                    kaihanShaName = tmpVal
                End If
            Next
        Next
    End If
    
    ' ���ʋL�^�V�[�g�Ɍ��ʂ�]�L����
    If rirekiNo <> "" Then
        targetSheet.Cells(i, columnName2Idx("G")).Value = Split(rirekiNo, "")
    End If
    
    If kaihanDataStr <> "" Then
        targetSheet.Cells(i, columnName2Idx("H")).Value = Split(kaihanDataStr, "")
    End If
    
    If kaihanShaName <> "" Then
        targetSheet.Cells(i, columnName2Idx("I")).Value = Split(kaihanShaName, "")
    End If

End Sub

' ���ŗ����V�[�g����Ő��Ɖ��ŎҖ����擾����
Sub parseKaihanRirekiSheet(targetSheet, i, parseTargetSheet)
    Dim lastRowNum As Integer
    Dim rirekiNo As String
    Dim kaihanDateStr As String
    Dim kaihanShaName As String
    
    lastRowNum = getLastIdxNo(parseTargetSheet, columnName2Idx("B"))

    ' Todo �J�n�s�A�擾�Ώۗ�����Ԃɍ��킹��
    rirekiNo = parseTargetSheet.Range("C" & lastRowNum).Text
    kaihanDateStr = parseTargetSheet.Range("D" & lastRowNum).Text
    kaihanShaName = parseTargetSheet.Range("E" & lastRowNum).Text
    
    Call DebugPrint(rirekiNo & "," & kaihanShaName)
    
    ' ���ʋL�^�V�[�g�Ɍ��ʂ�]�L����
    If rirekiNo <> "" Then
        targetSheet.Cells(i, columnName2Idx("G")).Value = Split(rirekiNo, "")
    End If
    
    If kaihanDateStr <> "" Then
        targetSheet.Cells(i, columnName2Idx("H")).Value = Split(kaihanDateStr, "")
    End If
    
    If kaihanShaName <> "" Then
        targetSheet.Cells(i, columnName2Idx("I")).Value = Split(kaihanShaName, "")
    End If
End Sub

' ���[�N�u�b�N��^����Ɖ��ŗ����V�[�g�̔ԍ���Ԃ��B�Ȃ����-1��Ԃ�
Function hasKaihanRirekiSheet(tmpBook As Workbook)
    Dim i As Integer
    
    For i = 1 To tmpBook.Worksheets.Count
        If tmpBook.Worksheets(i).Name Like "���ŗ���" Then
            hasKaihanRirekiSheet = i
            Exit Function
        End If
    Next

    hasKaihanRirekiSheet = -1
End Function

' �p�X����t�@�C�������擾
Function getFileName(filePath)
    Dim fileName As String

    fileName = FSO.getFileName(filePath)

    getFileName = fileName
End Function

' �p�X����g���q���擾
Function getExtentionName(filePath)
    Dim extentionName As String

    extentionName = FSO.getExtensionName(filePath)

    getExtentionName = extentionName
End Function

' Debug.Print�̃��b�p�[
Sub DebugPrint(msg)
    If IS_DEBUG = True Then
        Debug.Print (msg)
    End If
End Sub

' �����̗�ԍ��̗�̍ŉ��s�ԍ����擾
Function getLastIdxNo(targetSheet, colNum)
    Dim lastIdxNo As Integer
    
    ' �����̗�ԍ��̗�̍ŉ��s�ԍ����擾
    lastIdxNo = targetSheet.Cells(Rows.Count, colNum).End(xlUp).Row

    getLastIdxNo = lastIdxNo
End Function

' ��ԍ�����
Function columnIdx2Name(ByVal colNum As Long) As String
    columnIdx2Name = Split(Columns(colNum).Address, "$")(2)
End Function

' �񖼁���ԍ�
Function columnName2Idx(ByVal colName As String) As Long
    columnName2Idx = Columns(colName).Column
End Function

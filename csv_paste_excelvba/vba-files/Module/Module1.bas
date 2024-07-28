Attribute VB_Name = "Module1"
Option Explicit

Sub Main()
    Dim pasted_path As String
    Dim ReturnBook As Workbook
    Dim TargetBook As Workbook
    Dim r As Integer
    Dim filePath As String
    Dim sheetNum As Integer
    Dim startAddr
    Dim tempVal
    Dim startTime
    Dim endTime
    
    startTime = Time
    Application.ScreenUpdating = False

    ' ���̃}�N�������s����Excel�t�@�C�����擾
    Set ReturnBook = ActiveWorkbook
    
    ' �e�p�����[�^�����W
    pasted_path = ActiveWorkbook.Sheets(1).Cells(3, 2).Value
    
    ' �ʃV�[�g���J���đ���
    Set TargetBook = Workbooks.Open(pasted_path)

    ' �\��t����t�@�C���̏����擾���邽�߁AReturnBook���A�N�e�B�u�ɂ���
    ReturnBook.Activate
    
    With ReturnBook.Sheets(1).Range("B6:D" & �Ō���A�h���X���擾����())
        For r = 1 To .Rows.Count
            filePath = ReturnBook.Sheets(1).Range(.Item(r, 1).Address(False, False)).Value
            sheetNum = ReturnBook.Sheets(1).Range(.Item(r, 2).Address(False, False)).Value
            startAddr = ReturnBook.Sheets(1).Range(.Item(r, 3).Address(False, False)).Value
            Call pasteSheet(TargetBook, filePath, sheetNum, startAddr)
        Next r
    End With
    
    TargetBook.Save
    TargetBook.Close
    Set TargetBook = Nothing
    
    Application.ScreenUpdating = True
    
    endTime = Time - startTime
    MsgBox "�I��" & vbCrLf & Minute(endTime) & ":" & Second(endTime)
End Sub

' �t�@�C�����J���Ďw�肵���V�[�g�Ɏw�肵���A�h���X����\��t����
Sub pasteSheet(TargetBook, filePath, sheetNum, startAddr)
    Debug.Print filePath & ":" & sheetNum & ":" & startAddr
    
    Dim buf As String
    Dim tmp As Variant
    Dim r As Integer
    Dim c As Integer
    Dim rng As Range
    
    Set rng = TargetBook.Sheets(sheetNum).Range(startAddr)
    Open filePath For Input As #1
        Do Until EOF(1)
            Line Input #1, buf
            tmp = Split(buf, ",")
            r = r + 1

            For c = 0 To UBound(tmp)
                With rng.Cells(r, c + 1)
                    .NumberFormat = "@"
                    .Value = tmp(c)
                End With
            Next c
        Loop
    Close #1
    
    Set rng = Nothing
End Sub

' A��̍Ō���A�h���X���擾����
Function �Ō���A�h���X���擾����()
    �Ō���A�h���X���擾���� = ActiveWorkbook.Sheets(1).Cells(Rows.Count, 1).End(xlUp).Row
End Function


Attribute VB_Name = "multisheetcreate"
Option Explicit

' �����V�[�g�ꊇ�쐬
Sub �����V�[�g�ꊇ�쐬()
    Dim �I���Z�� As Range
    Dim i as Integer
    Dim sheetFound As Boolean
    Dim currentSheet as Worksheet

    set currentSheet = ActiveWorkBook.ActiveSheet

    i = Activeworkbook.Sheets.Count

    For Each �I���Z�� In Selection
        Debug.Print �I���Z��.Value

        ' �����V�[�g�����݂��Ȃ�������쐬����
        sheetFound = �V�[�g����(�I���Z��.Value)

        If sheetFound = False Then
            ActiveWorkBook.Sheets.Add(after:=ActiveWorkBook.Sheets(i)).Name = �I���Z��.Value
            i = i + 1
        Else
            Debug.Print �I���Z��.Value & "�V�[�g�͊��ɑ��݂��܂��B"
        End If
    Next �I���Z��

    ' ���̃V�[�g�ɖ߂�
    currentSheet.Activate
End Sub

' �V�[�g���݊m�F
Function �V�[�g����(sheetName As String)
    Dim ws As Worksheet
    
    ' �G���[�����Ă����s
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    ' �G���[����������
    On Error GoTo 0
    
    ' �V�[�g�����݂�����True���Ԃ�
    �V�[�g���� = Not ws Is Nothing
End Function

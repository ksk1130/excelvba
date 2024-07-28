Attribute VB_Name = "Module1"
Option Explicit

Sub Main()
    Dim EndRowNum As Integer
    Dim WorkSheetName As String
    Dim ws As Worksheet
    Dim i As Integer

    EndRowNum = GetEndRowNum("C")
    Debug.Print EndRowNum

    WorkSheetName = "work"

    ' work�V�[�g�����݂���ꍇ�͍폜���A�V�K�쐬
    For i = 1 To ThisWorkbook.Sheets.Count
        If ThisWorkbook.Sheets(i).Name = WorkSheetName Then
            Application.DisplayAlerts = False
            ThisWorkbook.Sheets(i).Delete
            Application.DisplayAlerts = True
            Exit For
        End If
    Next i
    Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    ws.Name = WorkSheetName
    
    ' �w�b�_�[�s�̒ǉ�
    ws.cells(1, 1).Value = "���t"
    ws.cells(1, 2).Value = "�S����"
    ws.cells(1, 3).Value = "�T�v"

    CreateWorkSheet EndRowNum, WorkSheetName

End Sub

' Sheet1�̃f�[�^��work�V�[�g�ɃR�s�[
Sub CreateWorkSheet(ByVal EndRowNum As Integer, ByVal WorkSheetName As String)
    Dim cn As Object
    Dim rs As Object
    Dim File_Name, Sql As String

    Dim CurRow As Integer
    File_Name = ThisWorkbook.FullName
    CurRow = 2

    Set cn = CreateObject("ADODB.Connection")
    Set rs = CreateObject("ADODB.Recordset")

    cn.Provider = "MSDASQL"
    cn.ConnectionString = "Driver={Microsoft Excel Driver (*.xls, *.xlsx, *.xlsm, *.xlsb)};" & "DBQ=" & File_Name & "; ReadOnly=True;"
    cn.Open

    Sql = "SELECT ���t,�S����,�T�v FROM [Sheet1$C2:E" & EndRowNum & "] "
    rs.Open Sql, cn, 3, 1  ' adOpenStatic = 3, adLockOptimistic = 1
    
    Do Until rs.EOF
        Sheets(WorkSheetName).Cells(CurRow, 1).Value = datevalue(rs!���t)
        Sheets(WorkSheetName).Cells(CurRow, 2).Value = rs!�S����
        Sheets(WorkSheetName).Cells(CurRow, 3).Value = rs!�T�v
        rs.MoveNext
        CurRow = CurRow + 1
    Loop
    
    rs.Close
    cn.Close

    Set rs = Nothing
    Set cn = Nothing
End Sub

' Sheet1�̍ŏI�s���擾    
Function GetEndRowNum(ByVal Column As String) As Integer
    Dim ws As Worksheet
    Dim LastRow As Integer

    ' �񖼂��ԗ�ԍ��ɕϊ�
    LastRow = columnName2Idx(Column)

    Set ws = ThisWorkbook.Sheets("Sheet1")

    GetEndRowNum = ws.Cells(ws.Rows.Count, LastRow).End(xlUp).Row
End Function

' �񖼁���ԍ�
Function columnName2Idx(ByVal colName As String) As Long
    columnName2Idx = Columns(colName).Column
End Function

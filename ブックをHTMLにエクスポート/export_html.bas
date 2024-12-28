Option Explicit

' ���[�N�u�b�N��HTML�`���ŕۑ�����
Sub export_html()

    ' ���s�m�F�̃_�C�A���O��\��
    If MsgBox("���̃}�N�������s���܂����H", vbYesNo) = vbNo Then
        Exit Sub
    End If

    ' �A�N�e�B�u���[�N�u�b�N�̃t�@�C�������擾
    Dim file_name As String
    file_name = ActiveWorkbook.Name
    Debug.Print file_name
    
    ' �t�@�C��������g���q���폜
    file_name = Left(file_name, InStrRev(file_name, ".") - 1)

    ' �t�@�C���ۑ��_�C�A���O��\��
    Dim html_name As String
    html_name = Application.GetSaveAsFilename( _
        InitialFileName:=file_name, _
        FileFilter:="HTML�t�@�C�� (*.htm), *.htm", _
        Title:="�ۑ�����w�肵�Ă�������")
    Debug.Print html_name

    ' �t�@�C���ۑ��_�C�A���O���L�����Z�����ꂽ�ꍇ�͏������I��
    If html_name = "False" Then
        Exit Sub
    End If

    ' �t�@�C����ۑ�
    ActiveWorkbook.SaveAs Filename:=html_name, FileFormat:=xlHtml, ReadOnlyRecommended:=False, CreateBackup:=False
End Sub

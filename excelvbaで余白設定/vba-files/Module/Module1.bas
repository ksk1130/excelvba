Attribute VB_Name = "Module1"
Option Explicit

Sub Hoge()
    Dim wb As Workbook

    ' �`�捂�����̂��߁A��ʍX�V���I�t�ɂ���
    Application.ScreenUpdating = False
    
    ' target.xlsx���J��
    Workbooks.Open Filename:="D:\Documents\github\excelvba\excelvba�ŗ]���ݒ�\target.xlsx"

    Set wb = Workbooks("target.xlsx")

    ' �V�[�g1��I��
    wb.Sheets(1).Select

    ' ����^�C�g���s��ݒ�
    wb.Sheets(1).PageSetup.PrintTitleRows = "$1:$4"

    ' �]����ݒ�(1.0cm�ɐݒ�)
    With wb.Sheets(1).PageSetup
        .LeftMargin = Application.centimetersToPoints(1)
        .RightMargin = Application.centimetersToPoints(1)
        .TopMargin = Application.centimetersToPoints(1)
        .BottomMargin = Application.centimetersToPoints(1)
        .HeaderMargin = Application.centimetersToPoints(1)
        .FooterMargin = Application.centimetersToPoints(1)
    End With

    ' target.xlsx��ۑ�
    wb.Save

    ' target.xlsx�����
    wb.Close

    Set wb = Nothing

    ' ��ʍX�V���I���ɂ���
    Application.ScreenUpdating = True
End Sub

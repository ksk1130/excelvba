Attribute VB_Name = "Module1"
Option Explicit

Sub Hoge()
    Dim wb As Workbook

    ' 描画高速化のため、画面更新をオフにする
    Application.ScreenUpdating = False
    
    ' target.xlsxを開く
    Workbooks.Open Filename:="D:\Documents\github\excelvba\excelvbaで余白設定\target.xlsx"

    Set wb = Workbooks("target.xlsx")

    ' シート1を選択
    wb.Sheets(1).Select

    ' 印刷タイトル行を設定
    wb.Sheets(1).PageSetup.PrintTitleRows = "$1:$4"

    ' 余白を設定(1.0cmに設定)
    With wb.Sheets(1).PageSetup
        .LeftMargin = Application.centimetersToPoints(1)
        .RightMargin = Application.centimetersToPoints(1)
        .TopMargin = Application.centimetersToPoints(1)
        .BottomMargin = Application.centimetersToPoints(1)
        .HeaderMargin = Application.centimetersToPoints(1)
        .FooterMargin = Application.centimetersToPoints(1)
    End With

    ' target.xlsxを保存
    wb.Save

    ' target.xlsxを閉じる
    wb.Close

    Set wb = Nothing

    ' 画面更新をオンにする
    Application.ScreenUpdating = True
End Sub

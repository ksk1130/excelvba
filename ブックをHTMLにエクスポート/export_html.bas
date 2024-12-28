Option Explicit

' ワークブックをHTML形式で保存する
Sub export_html()

    ' 実行確認のダイアログを表示
    If MsgBox("このマクロを実行しますか？", vbYesNo) = vbNo Then
        Exit Sub
    End If

    ' アクティブワークブックのファイル名を取得
    Dim file_name As String
    file_name = ActiveWorkbook.Name
    Debug.Print file_name
    
    ' ファイル名から拡張子を削除
    file_name = Left(file_name, InStrRev(file_name, ".") - 1)

    ' ファイル保存ダイアログを表示
    Dim html_name As String
    html_name = Application.GetSaveAsFilename( _
        InitialFileName:=file_name, _
        FileFilter:="HTMLファイル (*.htm), *.htm", _
        Title:="保存先を指定してください")
    Debug.Print html_name

    ' ファイル保存ダイアログがキャンセルされた場合は処理を終了
    If html_name = "False" Then
        Exit Sub
    End If

    ' ファイルを保存
    ActiveWorkbook.SaveAs Filename:=html_name, FileFormat:=xlHtml, ReadOnlyRecommended:=False, CreateBackup:=False
End Sub

Attribute VB_Name = "multisheetcreate"
Option Explicit

' 複数シート一括作成
Sub 複数シート一括作成()
    Dim 選択セル As Range
    Dim i as Integer
    Dim sheetFound As Boolean
    Dim currentSheet as Worksheet

    set currentSheet = ActiveWorkBook.ActiveSheet

    i = Activeworkbook.Sheets.Count

    For Each 選択セル In Selection
        Debug.Print 選択セル.Value

        ' 同名シートが存在しなかったら作成する
        sheetFound = シート存否(選択セル.Value)

        If sheetFound = False Then
            ActiveWorkBook.Sheets.Add(after:=ActiveWorkBook.Sheets(i)).Name = 選択セル.Value
            i = i + 1
        Else
            Debug.Print 選択セル.Value & "シートは既に存在します。"
        End If
    Next 選択セル

    ' 元のシートに戻る
    currentSheet.Activate
End Sub

' シート存在確認
Function シート存否(sheetName As String)
    Dim ws As Worksheet
    
    ' エラー合っても続行
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    ' エラー処理無効化
    On Error GoTo 0
    
    ' シートが存在したらTrueが返る
    シート存否 = Not ws Is Nothing
End Function

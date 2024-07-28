Attribute VB_Name = "Module1"
Option Explicit

Sub ボタン1_Click()
    Dim targetPath As String
    
    ' ファイル出力先
    targetPath = ActiveWorkbook.Worksheets(1).Cells(2, 3).Value
    
    Call export_file(ActiveWorkbook.Worksheets(2), targetPath)
    
    Debug.Print "End"

End Sub

' シートの内容をCSV形式(UTF-8, LF)で出力する
Sub export_file(targetWorksheet, targetPath)
    Dim maxRowNum As Long
    Dim targetRange As Range
    Dim i As Long
    Dim sheetName As String
    Dim fw As Variant
    Dim byteData() As Byte

    sheetName = targetWorksheet.Name

    ' 最下行の行番号を取得
    maxRowNum = getMaxRowNum(targetWorksheet)
    
    Set fw = CreateObject("ADODB.Stream")
    fw.Charset = "UTF-8"
    fw.Open
    
    ' ファイル出力対象範囲を決定
    Set targetRange = targetWorksheet.Range("A1:B" & maxRowNum)
    
    ' A1B1A2B2...の順で走査
    For i = 1 To targetRange.Count
        
        ' B列を処理後に改行コードを付与
        If i Mod 2 = 0 Then
            fw.WriteText targetRange.Item(i) & vbLf, 0
        Else
            fw.WriteText targetRange.Item(i) & ",", 0
        End If
        
    Next
    
    ' BOMなしUTF8作成のための作業
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

' A列の最下行番号を取得する
Function getMaxRowNum(targetWorksheet)
    Dim maxRowNum As Long
    
    maxRowNum = targetWorksheet.Cells(Rows.Count, 1).End(xlUp).Row

    getMaxRowNum = maxRowNum
End Function

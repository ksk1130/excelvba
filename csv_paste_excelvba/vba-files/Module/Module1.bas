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

    ' このマクロを実行したExcelファイルを取得
    Set ReturnBook = ActiveWorkbook
    
    ' 各パラメータを収集
    pasted_path = ActiveWorkbook.Sheets(1).Cells(3, 2).Value
    
    ' 別シートを開いて操作
    Set TargetBook = Workbooks.Open(pasted_path)

    ' 貼り付けるファイルの情報を取得するため、ReturnBookをアクティブにする
    ReturnBook.Activate
    
    With ReturnBook.Sheets(1).Range("B6:D" & 最後尾アドレスを取得する())
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
    MsgBox "終了" & vbCrLf & Minute(endTime) & ":" & Second(endTime)
End Sub

' ファイルを開いて指定したシートに指定したアドレスから貼り付ける
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

' A列の最後尾アドレスを取得する
Function 最後尾アドレスを取得する()
    最後尾アドレスを取得する = ActiveWorkbook.Sheets(1).Cells(Rows.Count, 1).End(xlUp).Row
End Function


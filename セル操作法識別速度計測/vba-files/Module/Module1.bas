Attribute VB_Name = "Module1"
Option Explicit

Const POWER As Integer = 3

Dim LOOP_COUNT As Long

' メイン処理
Sub Hoge()    
    ' 画面描画をオフにする
    Application.ScreenUpdating = False
    
    ' "copy"シートが存在していれば削除
    ' "result"シートが存在していれば削除
    On Error Resume Next
    Application.DisplayAlerts = False
    Sheets("copy").Delete
    Sheets("result").Delete
    Application.DisplayAlerts = True
    
    ' "result"シートを作成
    Sheets.Add(After:=Sheets(Sheets.Count)).Name = "result"
    Dim sheet_result As Worksheet
    Set sheet_result = Sheets("result")
    sheet_result.Cells(1, 1).Value = "処理内容"
    sheet_result.Cells(2, 1).Value = "セルを使ったループ処理"
    sheet_result.Cells(3, 1).Value = "1次元配列を使ったループ処理"
    sheet_result.Cells(4, 1).Value = "2次元配列を使ったループ処理"
    sheet_result.Cells(1, 2).Value = "値設定処理時間"
    sheet_result.Cells(1, 3).Value = "コピー処理時間"
    
    ' "copy"シートを作成
    Sheets.Add(After:=Sheets(Sheets.Count)).Name = "copy"
    Dim copy_sheet As Worksheet
    Set copy_sheet = Sheets("copy")

    Dim value_sheet As Worksheet
    Set value_sheet = Sheets("value")
    
    ' 10,100,1000回のループ処理を行う
    Dim i as Integer
    For i = 1 To POWER
        LOOP_COUNT = 10 ^ i        

        ' 全セルをクリア
        value_sheet.Cells.ClearContents
        copy_sheet.Cells.ClearContents
    
        Call ForLoopCells(value_sheet, copy_sheet, sheet_result)
    
        Call ForLoopRows(value_sheet, copy_sheet, sheet_result)
    
        Call ForLoopArrays(value_sheet, copy_sheet, sheet_result)
    Next i

    ' 画面描画をオンにする
    Application.ScreenUpdating = True

    MsgBox "処理が完了しました。"

End Sub
    
' セルを使ったループ処理
Sub ForLoopCells(value_sheet as Worksheet, copy_sheet as Worksheet, sheet_result as Worksheet)
    Dim i As Long
    Dim j As Long
    Dim colName As String
    Dim startTime As Double
    Dim endTime As Double
    Dim elapsedTime As Double

    ' valueシートに値を設定
    startTime = Timer
    For i = 1 To LOOP_COUNT
        For j = 1 To LOOP_COUNT
            value_sheet.Cells(i, j).Value = i * j
        Next j
    Next i
    endTime = Timer
    elapsedTime = endTime - startTime
    sheet_result.Cells(2, 2).Value = elapsedTime

    ' copyシートに値をコピー(セルを使ったループ処理)
    startTime = Timer
    For i = 1 To LOOP_COUNT
        For j = 1 To LOOP_COUNT
            copy_sheet.Cells(i, j).Value = value_sheet.Cells(i, j).Value
        Next j
    Next i
    endTime = Timer
    elapsedTime = endTime - startTime
    sheet_result.Cells(2, 3).Value = elapsedTime
End Sub

' 1次元配列を使ったループ処理
Sub ForLoopRows(value_sheet as Worksheet, copy_sheet as Worksheet, sheet_result as Worksheet)
    Dim i As Long
    Dim j As Long
    Dim colName As String
    Dim startTime As Double
    Dim endTime As Double
    Dim elapsedTime As Double

    ' 1次元配列を作成
    Dim arr1 As Variant

    ' jの値(列番号)を列名に変換
    colName = ColNumToColName(LOOP_COUNT, value_sheet)
    
    ' valueシートに値を設定
    startTime = Timer
    For i = 1 To LOOP_COUNT
        ' 配列の初期化
        ReDim arr1(LOOP_COUNT - 1)
        For j = 1 To LOOP_COUNT
            arr1(j - 1) = i * j
        Next j

        ' A1を起点に出力
        value_sheet.Range("A" & i & ":" & colName & i).Value = arr1
    Next i
    endTime = Timer
    elapsedTime = endTime - startTime
    sheet_result.Cells(3, 2).Value = elapsedTime

    ' copyシートに値をコピー(1次元配列を使ったループ処理)
    startTime = Timer
    For i = 1 To LOOP_COUNT
        arr1 = value_sheet.Range("A" & i & ":" & colName & i).Value
        copy_sheet.Range("A" & i & ":" & colName & i).Value = arr1
    Next i
    endTime = Timer
    elapsedTime = endTime - startTime
    sheet_result.Cells(3, 3).Value = elapsedTime
End Sub

' 2次元配列を使ったループ処理
Sub ForLoopArrays(value_sheet as Worksheet, copy_sheet as Worksheet, sheet_result as Worksheet)
    Dim i As Long
    Dim j As Long
    Dim colName As String
    Dim startTime As Double
    Dim endTime As Double
    Dim elapsedTime As Double

    ' 2次元配列を作成
    Dim arr2 As Variant

    ' jの値(列番号)を列名に変換
    colName = ColNumToColName(LOOP_COUNT, value_sheet)
    
    ' 配列の初期化
    ReDim arr2(LOOP_COUNT - 1, LOOP_COUNT - 1)
    
    ' valueシートに値を設定
    startTime = Timer
    For i = 1 To LOOP_COUNT
        For j = 1 To LOOP_COUNT
            arr2(i - 1, j - 1) = i * j
        Next j
    Next i    
    ' arr2をA1を起点に出力
    value_sheet.Range("A1:" & colName & LOOP_COUNT).Value = arr2
    endTime = Timer
    elapsedTime = endTime - startTime
    sheet_result.Cells(4, 2).Value = elapsedTime

    ' copyシートに値をコピー
    startTime = Timer
    arr2 = value_sheet.Range("A1:" & colName & LOOP_COUNT).Value
    copy_sheet.Range("A1:" & colName & LOOP_COUNT).Value = arr2
    endTime = Timer
    elapsedTime = endTime - startTime
    sheet_result.Cells(4, 3).Value = elapsedTime
End Sub

' 列番号を列名に変換
Function ColNumToColName(colNum As Long, value_sheet as Worksheet) As String
    Dim colName As String
    colName = Split(value_sheet.Cells(1, colNum).Address, "$")(1)
    ColNumToColName = colName
End Function

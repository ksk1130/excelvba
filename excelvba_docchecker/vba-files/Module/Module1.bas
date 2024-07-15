Attribute VB_Name = "Module1"
Option Explicit

' 定数定義
' TrueにするとDebug.Printが有効
Const IS_DEBUG As Boolean = True

' 共通変数定義
' FSOは共通定義を利用
Dim FSO As Object

Sub docchecker()

    Dim rootDir As String
    Dim sheetName As String

    ' オブジェクトはプロシージャ内での宣言が必要なためここで実施
    Set FSO = CreateObject("Scripting.FileSystemObject")

    ' フォルダ選択ダイアログ
    With Application.FileDialog(msoFileDialogFolderPicker)
        If .Show = True Then
            rootDir = .SelectedItems(1)
        End If
    End With

    Call DebugPrint(rootDir)

    ' フォルダパス有り無しチェック
    If rootDir = "" Then
        MsgBox "起点となるフォルダを選択してください。"
        Exit Sub
    End If

    ' 結果出力用シートの追加
    sheetName = addResultSheet()

    ' フォルダ配下のファイルを走査
    Call getFilesRecursive(rootDir, sheetName)
    
    ' 出力結果シートを基にチェックを実施
    Call checkDocument(sheetName)
    
    MsgBox ("終了")

End Sub

' 結果出力用シートの追加
Function addResultSheet()
    Dim NewWorkSheet As Worksheet
    Dim sheetName As String

    ' シート名は年月日時分秒
    sheetName = Format(Now, "yyyymmddhhnnss")

    Set NewWorkSheet = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    NewWorkSheet.Name = sheetName
    
    ' ヘッダ行の追加
    NewWorkSheet.Cells(2, columnName2Idx("B")).Value = "No"
    NewWorkSheet.Cells(2, columnName2Idx("C")).Value = "FilePath"
    NewWorkSheet.Cells(2, columnName2Idx("D")).Value = "Extention"
    NewWorkSheet.Cells(2, columnName2Idx("E")).Value = "FileName"
    NewWorkSheet.Cells(2, columnName2Idx("F")).Value = "改版履歴シートあり"
    NewWorkSheet.Cells(2, columnName2Idx("G")).Value = "版数"
    NewWorkSheet.Cells(2, columnName2Idx("H")).Value = "改版日"
    NewWorkSheet.Cells(2, columnName2Idx("I")).Value = "作成者名"
    NewWorkSheet.Cells(2, columnName2Idx("J")).Value = "ファイル名に版数あり"
    NewWorkSheet.Cells(2, columnName2Idx("K")).Value = "ファイル名に日付あり"
    NewWorkSheet.Cells(2, columnName2Idx("L")).Value = "新ファイル名"

    addResultSheet = sheetName
End Function

' 引数で与えられたパスを再帰的に走査する
Sub getFilesRecursive(path, sheetName)
    Dim objFolder As Object
    Dim objFile As Object

    ' GetFolder(フォルダ名).SubFoldersでフォルダ配下のフォルダ一覧を取得
    For Each objFolder In FSO.GetFolder(path).SubFolders
        Call getFilesRecursive(objFolder.path, sheetName)
    Next

    ' GetFolder(フォルダ名).Filesでフォルダ配下のファイル一覧を取得
    ' 結果をシートに記録
    For Each objFile In FSO.GetFolder(path).Files
        Call writeFilePathToSheet(objFile, sheetName)
    Next

End Sub

' 走査結果をシートに転記する
Sub writeFilePathToSheet(objFile, sheetName)
    Dim targetSheet As Worksheet
    
    Dim idx As Integer
    Dim idxNo As Integer

    Set targetSheet = ThisWorkbook.Worksheets(sheetName)
    
    ' idx = Excel行番号 , idxNo = 通し番号
    idx = getLastIdxNo(targetSheet, columnName2Idx("B")) + 1
    idxNo = idx - 2
    
    ' 結果シートには以下を転記する
    ' ファイルパス
    ' 拡張子
    ' ファイル名
    targetSheet.Cells(idx, columnName2Idx("B")).Value = idxNo
    targetSheet.Cells(idx, columnName2Idx("C")).Value = objFile.path
    targetSheet.Cells(idx, columnName2Idx("D")).Value = getExtentionName(objFile.path)
    targetSheet.Cells(idx, columnName2Idx("E")).Value = getFileName(objFile.path)
End Sub

' 個々のドキュメントについて確認を行う
Sub checkDocument(sheetName)
    Dim targetSheet As Worksheet
    Dim lastRowNum As Integer
    Dim i As Integer
    Dim filePath As String
    Dim extention As String
    Dim fileName As String
    Dim tmpBook As Workbook
    Dim kaihanRirekiSheetNum As Integer
    Dim newFileName As String

    ' Word操作用の定義
    Dim wordApp As Object
    Dim wordDoc As Object
    
    Set targetSheet = ThisWorkbook.Worksheets(sheetName)
    lastRowNum = getLastIdxNo(targetSheet, columnName2Idx("B"))
    Call DebugPrint(lastRowNum)
    
    ' ヘッダ行の1行下から、最下行まで走査
    For i = 3 To lastRowNum
        filePath = targetSheet.Cells(i, columnName2Idx("C")).Value
        extention = targetSheet.Cells(i, columnName2Idx("D")).Value
        fileName = targetSheet.Cells(i, columnName2Idx("E")).Value
    
        Call DebugPrint(extention)

        ' 拡張子チェック
        ' Wordの場合
        If extention Like "doc*" Then
            'Word文書の場合は都度Wordの起動停止を行う
            Set wordApp = CreateObject("Word.Application")
            wordApp.Visible = True
  
            Set wordDoc = wordApp.Documents.Open(fileName:=filePath, ReadOnly:=True)

            Call parseWordDoc(targetSheet, i, wordDoc)
            
            wordDoc.Close (False)
            wordApp.Quit

        ' Excelの場合
        ElseIf extention Like "xls*" Then
            Call DebugPrint("xls")
        
            Set tmpBook = Workbooks.Open(fileName:=filePath, ReadOnly:=True)
            
            kaihanRirekiSheetNum = hasKaihanRirekiSheet(tmpBook)
            
            ' 改版履歴シートあり
            If kaihanRirekiSheetNum > 0 Then
                targetSheet.Cells(i, columnName2Idx("F")).Value = "あり"
                
                ' 改版履歴シートから版数と記入者を取得して結果を転記
                Call parseKaihanRirekiSheet(targetSheet, i, tmpBook.Worksheets(kaihanRirekiSheetNum))
            
            ' 改版履歴シートなし -> 何もしない
            Else
                targetSheet.Cells(i, columnName2Idx("F")).Value = "なし"
            End If
            
            tmpBook.Close (False)
        Else
            Call DebugPrint("Other:" & extention)
        End If
 
        ' リネーム
        ' 新ファイル名を生成
        newFileName = getNewFileName(fileName, targetSheet, i)
    
    Next
    
End Sub

' ファイル名に版数と日付が含まれるかチェックし、新ファイル名を生成
Function getNewFileName(fileName As String, targetSheet As Worksheet, rowIdx As Integer)
    Dim tmpArr As Variant
    Dim arrLen As Integer
    Dim fileNameWithoutExtention As String
    Dim rirekiNo As String
    Dim kaihanDateStr As String
    Dim i As Integer
    Dim newFileName As String
    
    ' ファイル名から拡張子を除く
    fileNameWithoutExtention = Mid(fileName, 1, InStrRev(fileName, ".") - 1)
    Call DebugPrint(fileNameWithoutExtention)

    ' "_"でファイル名を分割して配列に格納
    tmpArr = Split(fileNameWithoutExtention, "_")
    
    ' 配列長を計算
    arrLen = UBound(tmpArr) - LBound(tmpArr) + 1
    
    ' 既定の書式ではない（分割後の配列長が4つ未満）場合は元のファイル名を返す
    If arrLen < 4 Then
        targetSheet.Cells(rowIdx, columnName2Idx("L")).Value = fileName
        getNewFileName = fileName
        Exit Function
    Else
        ' 最後から2つ目：版数、最後：改版日付として値を取得
        rirekiNo = tmpArr(arrLen - 2)
        kaihanDateStr = tmpArr(arrLen - 1)
        
        Call DebugPrint(rirekiNo & "," & kaihanDateStr)
        
        '　版数、改版日付ともに日付の場合は新ファイル名を生成して返却
        If IsNumeric(rirekiNo) = True And IsNumeric(kaihanDateStr) = True Then
            targetSheet.Cells(rowIdx, columnName2Idx("J")).Value = Split(rirekiNo, "")
            targetSheet.Cells(rowIdx, columnName2Idx("K")).Value = Split(kaihanDateStr, "")
            
            ' 版数、改版日は改版履歴の値を基に生成。改版履歴に値がなければ元の値で生成
            If targetSheet.Cells(rowIdx, columnName2Idx("G")).Value = "" Or targetSheet.Cells(rowIdx, columnName2Idx("H")).Value = "" Then
                newFileName = fileName
                targetSheet.Cells(rowIdx, columnName2Idx("L")).Value = newFileName
                getNewFileName = newFileName
                Exit Function
            Else
                ' 版数、改版日以外は元の値で生成
                For i = 0 To arrLen - 3
                    newFileName = newFileName & tmpArr(i) & "_"
                Next
            
                ' 版数
                newFileName = newFileName & targetSheet.Cells(rowIdx, columnName2Idx("G")).Value & "_"
                
                ' 改版日付
                newFileName = newFileName & Format(targetSheet.Cells(rowIdx, columnName2Idx("H")).Value, "yyyymmdd") & "_"
                
                ' 拡張子
                newFileName = newFileName & "." & targetSheet.Cells(rowIdx, columnName2Idx("D")).Value
            
                targetSheet.Cells(rowIdx, columnName2Idx("L")).Value = newFileName
            
                getNewFileName = newFileName
                Exit Function
            End If
        End If
    End If
End Function

' Word文書内から版数と改版者名を取得する
Sub parseWordDoc(targetSheet, i, wordDoc)
    Dim rirekiNo As String
    Dim kaihanDataStr As String
    Dim kaihanShaName As String
    Dim tbl As Object
    Dim j As Integer
    Dim k As Integer
    Dim tmpVal As String
    
    ' 文書内の表の数をチェックする。0なら何もしない
    If wordDoc.Tables.Count < 1 Then
        rirekiNo = "表なし"
        kaihanShaName = "表なし"
    Else
        ' 表がある場合は1つ目の表を走査する
        Set tbl = wordDoc.Tables(1)
        
        For j = 1 To tbl.Rows.Count
            For k = 1 To tbl.Columns.Count
                tmpVal = Left(tbl.Cell(j, k), Len(tbl.Cell(j, k)) - 2)
            
                If tmpVal <> "" And k = 1 Then
                    rirekiNo = tmpVal
                ElseIf tmpVal <> "" And k = 2 Then
                    kaihanDataStr = tmpVal
                ElseIf tmpVal <> "" And k = 3 Then
                    kaihanShaName = tmpVal
                End If
            Next
        Next
    End If
    
    ' 結果記録シートに結果を転記する
    If rirekiNo <> "" Then
        targetSheet.Cells(i, columnName2Idx("G")).Value = Split(rirekiNo, "")
    End If
    
    If kaihanDataStr <> "" Then
        targetSheet.Cells(i, columnName2Idx("H")).Value = Split(kaihanDataStr, "")
    End If
    
    If kaihanShaName <> "" Then
        targetSheet.Cells(i, columnName2Idx("I")).Value = Split(kaihanShaName, "")
    End If

End Sub

' 改版履歴シートから版数と改版者名を取得する
Sub parseKaihanRirekiSheet(targetSheet, i, parseTargetSheet)
    Dim lastRowNum As Integer
    Dim rirekiNo As String
    Dim kaihanDateStr As String
    Dim kaihanShaName As String
    
    lastRowNum = getLastIdxNo(parseTargetSheet, columnName2Idx("B"))

    ' Todo 開始行、取得対象列を実態に合わせる
    rirekiNo = parseTargetSheet.Range("C" & lastRowNum).Text
    kaihanDateStr = parseTargetSheet.Range("D" & lastRowNum).Text
    kaihanShaName = parseTargetSheet.Range("E" & lastRowNum).Text
    
    Call DebugPrint(rirekiNo & "," & kaihanShaName)
    
    ' 結果記録シートに結果を転記する
    If rirekiNo <> "" Then
        targetSheet.Cells(i, columnName2Idx("G")).Value = Split(rirekiNo, "")
    End If
    
    If kaihanDateStr <> "" Then
        targetSheet.Cells(i, columnName2Idx("H")).Value = Split(kaihanDateStr, "")
    End If
    
    If kaihanShaName <> "" Then
        targetSheet.Cells(i, columnName2Idx("I")).Value = Split(kaihanShaName, "")
    End If
End Sub

' ワークブックを与えると改版履歴シートの番号を返す。なければ-1を返す
Function hasKaihanRirekiSheet(tmpBook As Workbook)
    Dim i As Integer
    
    For i = 1 To tmpBook.Worksheets.Count
        If tmpBook.Worksheets(i).Name Like "改版履歴" Then
            hasKaihanRirekiSheet = i
            Exit Function
        End If
    Next

    hasKaihanRirekiSheet = -1
End Function

' パスからファイル名を取得
Function getFileName(filePath)
    Dim fileName As String

    fileName = FSO.getFileName(filePath)

    getFileName = fileName
End Function

' パスから拡張子を取得
Function getExtentionName(filePath)
    Dim extentionName As String

    extentionName = FSO.getExtensionName(filePath)

    getExtentionName = extentionName
End Function

' Debug.Printのラッパー
Sub DebugPrint(msg)
    If IS_DEBUG = True Then
        Debug.Print (msg)
    End If
End Sub

' 引数の列番号の列の最下行番号を取得
Function getLastIdxNo(targetSheet, colNum)
    Dim lastIdxNo As Integer
    
    ' 引数の列番号の列の最下行番号を取得
    lastIdxNo = targetSheet.Cells(Rows.Count, colNum).End(xlUp).Row

    getLastIdxNo = lastIdxNo
End Function

' 列番号→列名
Function columnIdx2Name(ByVal colNum As Long) As String
    columnIdx2Name = Split(Columns(colNum).Address, "$")(2)
End Function

' 列名→列番号
Function columnName2Idx(ByVal colName As String) As Long
    columnName2Idx = Columns(colName).Column
End Function

Attribute VB_Name = "Module1"
Option Explicit

' 未達
' 年号のはざま(1989/1/6とか)が次の元号になってしまっている
Sub Macro1()
Attribute Macro1.VB_ProcData.VB_Invoke_Func = " \n14"
    Dim i As Integer
    Dim j As Integer
    Dim outputPath As String
    
    If MsgBox("処理を実行しますか？", vbYesNo) = vbNo Then
        Exit Sub
    End If
    
    outputPath = ActiveSheet.Cells(4, 2).Value
    outputPath = outputPath & "\wareki_all.txt"
 
    Open outputPath For Output As #1
        
    For i = 1868 To 2100
        For j = 1 To 12
            If i >= 1868 And i < 1912 Then '明治
                ' 1868/1-9はスキップ
                If i = 1868 Then '明治
                    Select Case j
                        Case Is >= 10
                            Print #1, "1" & Format(i - 1867, "00") & Format(j, "00")
                    End Select
                End If
            
                Print #1, "1" & Format(i - 1867, "00") & Format(j, "00")
            ElseIf i >= 1912 And i <= 1926 Then '大正
                If i = 1912 Then '大正
                    Select Case j
                        Case Is < 7
                            Print #1, "1" & Format(i - 1867, "00") & Format(j, "00")
                        Case Is >= 7
                            Print #1, "2" & Format(i - 1911, "00") & Format(j, "00")
                    End Select
                ElseIf i = 1926 Then
                    Select Case j
                        Case Is < 12
                            Print #1, "2" & Format(i - 1911, "00") & Format(j, "00")
                        Case Else
                            Print #1, "3" & Format(i - 1925, "00") & Format(j, "00")
                    End Select
                Else
                   Print #1, "2" & Format(i - 1911, "00") & Format(j, "00")
                End If
            ElseIf i > 1926 And i < 1989 Then '昭和
                Print #1, "3" & Format(i - 1925, "00") & Format(j, "00")
            ElseIf i >= 1989 And i <= 2019 Then '平成
                If i = 2019 Then
                    Select Case j
                        Case Is < 5
                            Print #1, "4" & Format(i - 1988, "00") & Format(j, "00")
                        Case Else
                            Print #1, "5" & Format(i - 2018, "00") & Format(j, "00")
                    End Select
                Else
                    Print #1, "4" & Format(i - 1988, "00") & Format(j, "00")
                End If
            Else ' 令和
                Print #1, "5" & Format(i - 2018, "00") & Format(j, "00")
            End If
        Next
    Next
    
    Close #1
    MsgBox ("処理終了")

End Sub

Option Explicit

'ファイル名の変更
Sub Rename()
    Dim column As Integer
    column = 9 'ファイル名の1行目
    
    If Range("$C$4").Value = "" Then '変更前パスチェック
        MsgBox "Fill the blank space.(C4)"
        Exit Sub
    End If
    If Range("$C$5").Value = "" Then '変更後パスチェック
        MsgBox "Fill the blank space.(C5)"
        Exit Sub
    End If
    If Range("$C$6").Value = "" Then '拡張子チェック
        MsgBox "Fill the blank space.(C6)"
        Exit Sub
    End If
    
    Do Until Cells(column, 2).Value = ""
        If Cells(column, 3).Value = "" Then 'ファイル名チェック
            MsgBox "Fill the blank space."
        Else
            FileCopy Range("$C$4").Value & "\" & Cells(column, 2).Value & "." & Range("$C$6").Value, _
                     Range("$C$5").Value & "\" & Cells(column, 3).Value & "." & Range("$C$6").Value
            Kill Range("$C$4").Value & "\" & Cells(column, 2).Value & "." & Range("$C$6").Value
        End If
        column = column + 1
    Loop
    
    MsgBox "名前の変更が完了しました。", vbInformation
End Sub
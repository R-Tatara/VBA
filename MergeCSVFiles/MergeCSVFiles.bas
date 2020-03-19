'複数のcsvをマージしてxlxs形式で保存
Sub MergeCSVFiles()
    Const NEW_NAME As String = "compilation.xlsx"
    Const NEW_DIR As String = "C:\Users\Ryosuke\Desktop\vba\" & NEW_NAME
    
    'ファイル選択ダイアログ
    selectDir = _
        Application.GetOpenFilename( _
            FileFilter:="CSVファイル(*.csv),*.csv", _
            FilterIndex:=1, _
            Title:="Choose file", _
            MultiSelect:=True _
        )
    
    If IsArray(selectDir) Then
        '新規ワークブックの作成
        Dim newWorkbook As Workbook
        
        Workbooks.Add
        Set newWorkbook = ActiveWorkbook
        newWorkbook.SaveAs Filename:=NEW_DIR, _
            FileFormat:=xlWorkbookDefault
        Workbooks(NEW_NAME).Sheets(1).Name = "Master"
        
        Dim i As Integer
        i = 1
        
        For Each eachDir In selectDir
            Dim newSheet As Worksheet
            Dim oldName As String
            '選択ファイルを開く
            Workbooks.Open eachDir
            oldName = Dir(eachDir)
            
            '新しいシートに追加
            Workbooks(oldName).Worksheets(1).Copy _
                After:=Workbooks(NEW_NAME).Worksheets(i)
            Set newSheet = newWorkbook.Worksheets(i + 1)
            
            '処理
            
            i = i + 1
            On Error Resume Next
            
            Application.DisplayAlerts = False
            Workbooks(oldName).Close
            Application.DisplayAlerts = True
        Next
        
        newWorkbook.Save
        MsgBox "Created ""compilation.xlsx"""
    End If
End Sub
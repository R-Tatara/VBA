'������csv���}�[�W����xlxs�`���ŕۑ�
Sub MergeCSVFiles()
    Const NEW_NAME As String = "compilation.xlsx"
    Const NEW_DIR As String = "C:\Users\Ryosuke\Desktop\vba\" & NEW_NAME
    
    '�t�@�C���I���_�C�A���O
    selectDir = _
        Application.GetOpenFilename( _
            FileFilter:="CSV�t�@�C��(*.csv),*.csv", _
            FilterIndex:=1, _
            Title:="Choose file", _
            MultiSelect:=True _
        )
    
    If IsArray(selectDir) Then
        '�V�K���[�N�u�b�N�̍쐬
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
            '�I���t�@�C�����J��
            Workbooks.Open eachDir
            oldName = Dir(eachDir)
            
            '�V�����V�[�g�ɒǉ�
            Workbooks(oldName).Worksheets(1).Copy _
                After:=Workbooks(NEW_NAME).Worksheets(i)
            Set newSheet = newWorkbook.Worksheets(i + 1)
            
            '����
            
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
Option Explicit

'Convert .csv into .xlsx
Sub ConvertCsvToXlsx()
    Dim fso As New FileSystemObject 'Need "Microsoft Scripting Runtime"
    Dim WorkbookName As String
    Dim OriginalName As String
    Dim Extension As String
    Dim i As Long

    '�g���q���������u�b�N���̎擾
    WorkbookName = ActiveWorkbook.FullName
    OriginalName = WorkbookName
    i = InStrRev(WorkbookName, ".")
    If i > 0 Then
        WorkbookName = Left(WorkbookName, i - 1)
    End If

    'csv�t�@�C���ɂ̂ݓK�p
    Extension = fso.GetExtensionName(ActiveWorkbook.Path & ActiveWorkbook.Name)
    If Extension <> "csv" Then
        MsgBox "csv�t�@�C���ɂ̂ݓK�p�\�ł�", vbExclamation
        End
    End If

    'xlsx�`���ŕۑ�
    On Error Resume Next
    ActiveWorkbook.SaveAs Filename:=WorkbookName & ".xlsx", _
                                      FileFormat:=xlOpenXMLWorkbook
    If Err.Number > 0 Then
        MsgBox "�t�@�C���̊g���q�͕ϊ�����܂���ł���"
        End
    End If

    'csv�t�@�C���̍폜
    Kill (OriginalName)

    '�������b�Z�[�W
    MsgBox "�t�@�C���̊g���q��ύX���܂���", vbInformation
End Sub


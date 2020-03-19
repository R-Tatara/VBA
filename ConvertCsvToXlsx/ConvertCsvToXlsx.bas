Option Explicit

'Convert .csv into .xlsx
Sub ConvertCsvToXlsx()
    Dim fso As New FileSystemObject 'Need "Microsoft Scripting Runtime"
    Dim WorkbookName As String
    Dim OriginalName As String
    Dim Extension As String
    Dim i As Long

    '拡張子を除いたブック名の取得
    WorkbookName = ActiveWorkbook.FullName
    OriginalName = WorkbookName
    i = InStrRev(WorkbookName, ".")
    If i > 0 Then
        WorkbookName = Left(WorkbookName, i - 1)
    End If

    'csvファイルにのみ適用
    Extension = fso.GetExtensionName(ActiveWorkbook.Path & ActiveWorkbook.Name)
    If Extension <> "csv" Then
        MsgBox "csvファイルにのみ適用可能です", vbExclamation
        End
    End If

    'xlsx形式で保存
    On Error Resume Next
    ActiveWorkbook.SaveAs Filename:=WorkbookName & ".xlsx", _
                                      FileFormat:=xlOpenXMLWorkbook
    If Err.Number > 0 Then
        MsgBox "ファイルの拡張子は変換されませんでした"
        End
    End If

    'csvファイルの削除
    Kill (OriginalName)

    '完了メッセージ
    MsgBox "ファイルの拡張子を変更しました", vbInformation
End Sub


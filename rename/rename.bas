Option Explicit

'�t�@�C�����̕ύX
Sub Rename()
    Dim column As Integer
    column = 9 '�t�@�C������1�s��
    
    If Range("$C$4").Value = "" Then '�ύX�O�p�X�`�F�b�N
        MsgBox "Fill the blank space.(C4)"
        Exit Sub
    End If
    If Range("$C$5").Value = "" Then '�ύX��p�X�`�F�b�N
        MsgBox "Fill the blank space.(C5)"
        Exit Sub
    End If
    If Range("$C$6").Value = "" Then '�g���q�`�F�b�N
        MsgBox "Fill the blank space.(C6)"
        Exit Sub
    End If
    
    Do Until Cells(column, 2).Value = ""
        If Cells(column, 3).Value = "" Then '�t�@�C�����`�F�b�N
            MsgBox "Fill the blank space."
        Else
            FileCopy Range("$C$4").Value & "\" & Cells(column, 2).Value & "." & Range("$C$6").Value, _
                     Range("$C$5").Value & "\" & Cells(column, 3).Value & "." & Range("$C$6").Value
            Kill Range("$C$4").Value & "\" & Cells(column, 2).Value & "." & Range("$C$6").Value
        End If
        column = column + 1
    Loop
    
    MsgBox "���O�̕ύX���������܂����B", vbInformation
End Sub
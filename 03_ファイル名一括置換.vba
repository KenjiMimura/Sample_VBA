' ----------------------------------------------------------------------
'   �t�@�C�����F�t�@�C�����ꊇ�u��
'   ������F	Excel
'   �@�\�����F  �w�肵���t�H���_���̑S�t�@�C���̃t�@�C�������w�肵�������Œu��
' ----------------------------------------------------------------------

'�萔��`
'���b�Z�[�W�֘A
Const DIR_TITLE As String = "�Ώۃt�H���_��I��"
Const MSG_SELECT As String = "�t�@�C������u���������t�@�C�����i�[���ꂽ�t�H���_���w�肵�ĉ������B" & vbCrLf & _
								"���@�w�肵���t�H���_���̑S�t�@�C���������ΏۂƂȂ�܂��B"
Const MSG_FIND As String = "�u���̑ΏۂƂȂ镶�������͂��ĉ������B"
Const MSG_REPLACE As String = "�u����̕��������͂��ĉ������B"
Const MSG_ERR As String = "�t�@�C�����Ƃ��Ďg�p�o���Ȃ��������܂܂�Ă��܂��B" & vbCrLf & "������x�A���͂��ĉ������B"
Const MSG_CONF As String = "���̓��e�Ńt�@�C�����̒u�������s���܂����A��낵���ł����H"
Const MSG_SUCCESS As String = "�t�@�C�����ύX����"
Const MSG_FAIL As String = "�t�@�C�����ύX���s"
Const MSG_SUSPEND As String = "�������I�����܂��B"
Const MSG_END As String = "�������������܂����B"
'���s�m�F��ʊ֘A
Const DLG_TITLE As String = "�m�F"
Const EXEC_HISTORY As String = "���s����"
Const TARGET_FOLDER = "�Ώۃt�H���_�@�F�@"
Const FIND_WORD = "�u���O������@�F�@"
Const REPLACE_WORD = "�u���㕶����@�F�@"
Const RESULT_HEAD As String = "���s����"
Const ERROR_HEAD As String = "�G���[���e"
Const PATH_HEAD As String = "�t�@�C���i�[�ꏊ"
Const BF_FILE_HEAD As String = "�ύX�O�t�@�C����"
Const AF_FILE_HEAD As String = "�ύX��t�@�C����"
'���̑��萔
Const HEADER_ROW As Integer = 1
Const FIRST_ROW As Integer = 2
Const RESULT_COL As String = "A"
Const AF_FILE_COL As String = "E"

'�񋓌^��`
'�w�b�_�[�֘A
Enum HEADER
    EXEC_RESULT = 1
    EXEC_ERROR
    FILE_PATH
    BF_FILE
    AF_FILE
End Enum

' ----------------------------------------------------------------------
'   �֐����FFileNameReplacement�i�又���j
'   �����F  �Ȃ�
'   �߂�l�F�Ȃ�
' ----------------------------------------------------------------------
Sub FileNameReplacement()
    '�ϐ��錾
    Dim Fso As Object
    Dim File As Object
    Dim NewSheet As Worksheet
    Dim FolderPath As String
    Dim NewFileName As String
    Dim FindStr As Variant
    Dim ReplaceStr As Variant
    Dim Result As Integer
    Dim SheetCnt As Integer
    Dim ExecCnt As Integer

    With Application
        '��ʍX�V��~
        .ScreenUpdating = False
        '�m�F���b�Z�[�W��\��
        .DisplayAlerts = False
    End With

    '�J�����g�f�B���N�g���ύX
    ChDir ThisWorkbook.Path

    '�Ώۂ̃t�H���_���w��
    MsgBox MSG_SELECT, vbInformation
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = DIR_TITLE                      '�^�C�g���̎w��
        .InitialFileName = ThisWorkbook.Path    '�����\���t�H���_�̎w��
        If .Show = True Then                    '�_�C�A���O��\�����Ė߂�l�𔻒�
            FolderPath = .SelectedItems(1)      '�t�H���_�̃p�X���擾
        Else
            Exit Sub
        End If
    End With
    
    '�������������
    If UserInput(FindStr, MSG_FIND) = False Then Exit Sub
    
    '�u���������
    If UserInput(ReplaceStr, MSG_REPLACE) = False Then Exit Sub
     
    '�����̎��s���m�F
    Result = MsgBox( _
                    TARGET_FOLDER & FolderPath & vbCrLf & vbCrLf & _
                    FIND_WORD & FindStr & vbCrLf & _
                    REPLACE_WORD & ReplaceStr & vbCrLf & vbCrLf & _
                    MSG_CONF, vbYesNo + vbQuestion, DLG_TITLE)
    '�����L�����Z����
    If Result = vbNo Then
        MsgBox MSG_SUSPEND
        Exit Sub
    End If

    '���s�����V�[�g�����݂��Ă���ꍇ�͍폜
    For SheetCnt = 1 To Worksheets.Count
        If Worksheets(SheetCnt).Name = EXEC_HISTORY Then
            Worksheets(EXEC_HISTORY).Delete
            Exit For
        End If
    Next SheetCnt
    
    '���s�����V�[�g���쐬
    Set NewSheet = ThisWorkbook.Worksheets.Add(after:=Worksheets(Worksheets.Count))
    NewSheet.Name = EXEC_HISTORY
    
    '���s�����V�[�g�̃w�b�_�[�s���쐬
    With NewSheet
        .Cells(HEADER_ROW, HEADER.EXEC_RESULT).Value = RESULT_HEAD
        .Cells(HEADER_ROW, HEADER.EXEC_ERROR).Value = ERROR_HEAD
        .Cells(HEADER_ROW, HEADER.FILE_PATH).Value = PATH_HEAD
        .Cells(HEADER_ROW, HEADER.BF_FILE).Value = BF_FILE_HEAD
        .Cells(HEADER_ROW, HEADER.AF_FILE).Value = AF_FILE_HEAD
    End With
    
    ExecCnt = FIRST_ROW
    Set Fso = CreateObject("Scripting.FileSystemObject")
    For Each File In Fso.getFolder(FolderPath).Files
        '�u����̃t�@�C�����𐶐�
        NewFileName = Replace(File.Name, FindStr, ReplaceStr)

        If File.Name <> NewFileName Then
            On Error Resume Next
            
            With NewSheet
                .Cells(ExecCnt, HEADER.FILE_PATH).Value = File.ParentFolder
                .Cells(ExecCnt, HEADER.BF_FILE).Value = File.Name
                .Cells(ExecCnt, HEADER.AF_FILE).Value = NewFileName
                
                '�t�@�C�����ύX
                Name File.Path As File.ParentFolder & "\" & NewFileName
                
                If Err.Number <> 0 Then
                    '�t�@�C�����̕ύX���o���Ȃ������ꍇ�̏���
                    .Cells(ExecCnt, HEADER.EXEC_RESULT).Value = MSG_FAIL
                    .Cells(ExecCnt, HEADER.EXEC_ERROR).Value = Err.Description
                    .Range(RESULT_COL & ExecCnt & ":" & AF_FILE_COL & ExecCnt).Font.Color = RGB(255, 0, 0) '�����F:��
                    Err.Clear
                Else
                    .Cells(ExecCnt, HEADER.EXEC_RESULT).Value = MSG_SUCCESS
                End If
            End With
            ExecCnt = ExecCnt + 1
        End If
    Next
    
    '�I�u�W�F�N�g���
    Set Fso = Nothing
    Set File = Nothing
    
    '�����������
    NewSheet.Columns(RESULT_COL & ":" & AF_FILE_COL).AutoFit

    With Application
        '��ʍX�V��~
        .ScreenUpdating = True
        '�m�F���b�Z�[�W��\��
        .DisplayAlerts = True
    End With

    '�I�����b�Z�[�W��\��
    MsgBox MSG_END, vbInformation
End Sub

' ----------------------------------------------------------------------
'   �֐����FUserInput�i���[�U�[���͏����j
'   ����1�F InputString�i���[�U�[�̓��͂��󂯎��ϐ��j
'   ����2�F OutputMessage�i�e�L�X�g�{�b�N�X�ɕ\������镶���j
'   �߂�l�FBollean�i����=True/�ُ�=False�j
' ----------------------------------------------------------------------
Function UserInput(InputString As Variant, OutputMessage As String) As Boolean
    '�ϐ��錾
    Dim CheckFlg As Boolean

    '�֐��̖߂�l��ݒ�
    UserInput = True
    
    '���[�U�[����
    Do
        CheckFlg = True
        InputString = InputBox(OutputMessage)
        If StrPtr(InputString) = 0 Then
            MsgBox MSG_SUSPEND
            UserInput = False
            Exit Function
        End If
        
        '�g�p�֎~�����̊m�F
        If InStr(InputString, "\") > 0 Then CheckFlg = False
        If InStr(InputString, "/") > 0 Then CheckFlg = False
        If InStr(InputString, ":") > 0 Then CheckFlg = False
        If InStr(InputString, "*") > 0 Then CheckFlg = False
        If InStr(InputString, "?") > 0 Then CheckFlg = False
        If InStr(InputString, """") > 0 Then CheckFlg = False
        If InStr(InputString, "<") > 0 Then CheckFlg = False
        If InStr(InputString, ">") > 0 Then CheckFlg = False
        If InStr(InputString, "|") > 0 Then CheckFlg = False
        
        '�g�p�֎~���������͂��ꂽ�ꍇ�̏���
        If CheckFlg = False Then
            MsgBox MSG_ERR
        End If
    Loop While CheckFlg = False
End Function

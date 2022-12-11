' ----------------------------------------------------------------------
'   �t�@�C�����FHTML�t�@�C���ǂݍ���
'   ������F	Access
'   �@�\�����F  �w�肵���t�H���_���̑SHTML�t�@�C�����珤�i�֘A�̃R�[�h���擾
'				�擾�������i�֘A�̃R�[�h�Ə��i����CSV�t�@�C����ACCESS�Ɏ�荞��
'				HTML�t�@�C���ɋL�ڂ���Ă������i�����G�N�Z���t�@�C���ŏo��
' ----------------------------------------------------------------------

'�萔��`
'�_�C�A���O�֘A
Const DLG_TITLE_FILE As String = "�t�@�C����I��"
Const DLG_TITLE_FOLDER As String = "�t�H���_��I��"
Const FILTER_DESC As String = "CSV�t�@�C��"
Const FILTER_EXT As String = "*.csv"
'���������֘A
Const EXT_HTM As String = "htm"
Const EXT_HTML As String = "html"
Const LINK_TAG As String = "href"
Const AMPERSAND As String = "&"
Const PRE_RELATIONAL_CODE As String = "&relational_code="
Const PRE_PRODUCT_CODE As String = "&product_code="
Const WORD_FORM_STR As String = "RELATIONAL_CODE,RELATIONAL_TYPE,PRODUCT_CODE,PRODUCT_NAME,BRANCH_NAME,TEAM_NAME,INFLOW_ROUTE"
Const WORD_FORM_INT As String = "QUANTITY,UNIT_PRICE,SALES,INFLOW_ROUTE_NO"
Const WORD_FORM_DATE As String = "RELEASE_DATE,BOOKING_DATE"
'�f�[�^�x�[�X�֘A
Const TABLE_PUBLISH_CODE As String = "PublishCode"
Const SQL_DELETE As String = "DELETE * FROM "
Const QUERY_C10 As String = "C10_ProductData"
'�t�@�C���֘A
Const COPY_FILE_NAME = "\Target.csv"
Const EXT_EXCEL As String = ".xlsx"
Const LINK_FILE_NAME As String = "DATA"
Const EXP_FILE_NAME As String = "�f�ڏ��i���"
Const EXP_DATA_NAME As String = "�f�ڏ��i�ꗗ"
'���b�Z�[�W�֘A
Const MSG_TITLE = "GetProductData"
Const MSG_END As String = "�������������܂���"
'�G���[�֘A
Const NOT_FOUND_HTML As String = "html�t�@�C����������܂���ł����B"
Const NOT_FOUND_STR_COL As String = "���i���t�@�C������Ώۂ̗񂪌�����܂���ł����B�i������j"
Const NOT_FOUND_INT_COL As String = "���i���t�@�C������Ώۂ̗񂪌�����܂���ł����B�i���l�j"
Const NOT_FOUND_DATE_COL As String = "���i���t�@�C������Ώۂ̗񂪌�����܂���ł����B�i���t�j"
'�����֘A
Const FORM_DATE8 As String = "yyyymmdd"
Const FORM_STR As String = "@"
Const FORM_INT As String = "0_ "
Const FORM_DATE As String = "yyyy/mm/dd"
'���l�֘A
Const ZOOM_POINT As Integer = 80
Const HEADER_ROW As Integer = 1
Const TOP_SHEET As Integer = 1

'�񋓌^��`
'�����N�������
Enum RESULT
    NONE = 0
    RELATIONAL_CODE
    PRODUCT_CODE
End Enum

'PublishCode�e�[�u���̃J����
Enum COL
    FILE_NAME = 0
    RELATIONAL_CODE
    PRODUCT_CODE
End Enum

'������ʃt���O
Enum PROC_FLG
    F_STR = 1
    F_INT
    F_DATE
End Enum

' ----------------------------------------------------------------------
'   �֐����FButton10_Click�i�t�@�C���I���{�^�������������j
'   �����F  �Ȃ�
'   �߂�l�F�Ȃ�
' ----------------------------------------------------------------------
Private Sub Button10_Click()
    '�t�@�C���I���_�C�A���O�{�b�N�X��\��
    With Application.FileDialog(msoFileDialogFilePicker)
        '�^�C�g����ݒ�
        .Title = DLG_TITLE_FILE
        
        '�t�@�C���̎�ނ�ݒ�
        .Filters.Clear
        .Filters.Add FILTER_DESC, FILTER_EXT
        .FilterIndex = 1
        
        '�����t�@�C���I���������Ȃ�
        .AllowMultiSelect = False
        
        '�����p�X��ݒ�
        If (Forms!F_StartUp!TextBox10.Value <> "") Then
            .InitialFileName = Forms!F_StartUp!TextBox10.Value
        Else
            .InitialFileName = CurrentProject.Path
        End If
        
        If (.Show <> 0) Then
            '�I�����ꂽ�t�@�C���̃p�X���e�L�X�g�{�b�N�X�ɐݒ�
            Forms!F_StartUp!TextBox10.Value = Trim(.SelectedItems.Item(1))
        Else
            Exit Sub
        End If
    End With
End Sub

' ----------------------------------------------------------------------
'   �֐����FButton20_Click�i�����J�n�{�^�������������j
'   �����F  �Ȃ�
'   �߂�l�F�Ȃ�
' ----------------------------------------------------------------------
Private Sub Button20_Click()
    '�ϐ��錾
    Dim CurPath As String
    Dim CurDate As String
    Dim TargetPath As String
    Dim ConvFilePath As String
    Dim ExpFilePath As String
    ReDim PublishCode(2, 0) As String
    
    '�V�X�e�����b�Z�[�W��\��
    DoCmd.SetWarnings False
    
    '�J�����g�t�H���_�[�̃p�X���擾
    CurPath = CurrentProject.Path
    
    '���݂̓��t��yyyymmdd�`���Ŏ擾
    CurDate = FORMAT(Date, FORM_DATE8)
    
    '�t�H���_�[�I���_�C�A���O�{�b�N�X��\��
    With Application.FileDialog(msoFileDialogFolderPicker)
        '�^�C�g����ݒ�
        .Title = DLG_TITLE_FOLDER
        
        '�����p�X��ݒ�
        .InitialFileName = CurPath
        
        If (.Show <> 0) Then
            '�I�����ꂽ�t�H���_�̃p�X���擾
            TargetPath = Trim(.SelectedItems.Item(1))
        Else
            Exit Sub
        End If
    End With
        
    'html�t�@�C��������
    If (FileSearch(TargetPath, PublishCode) = False) Then
        MsgBox NOT_FOUND_HTML, vbOKOnly, MSG_TITLE
        Exit Sub
    End If
    
    'html���̑ΏۃR�[�h�����[�J���e�[�u���Ƃ��č쐬
    MakePublishData PublishCode
    
    '���i�̏����擾
    GetProductData CurPath, CurDate, ConvFilePath
    
    '�G�N�Z���t�@�C���ɏo��
    If (ExportBook(CurPath, CurDate, ConvFilePath, ExpFilePath) = False) Then Exit Sub
    
    '�o�͂����G�N�Z���t�@�C����ҏW
    EditBook ExpFilePath
    
    '�V�X�e�����b�Z�[�W�\��
    DoCmd.SetWarnings True
    
    '�����I�����b�Z�[�W
    MsgBox MSG_END, vbOKOnly, MSG_TITLE
End Sub

' ----------------------------------------------------------------------
'   �֐����FFileSearch�i�t�@�C�����������j
'   ����1�F TargetPath�i���[�U�[���I�������t�H���_�[�̃p�X�j
'   ����2�F PublishCode�ihtml�t�@�C�����̎擾�ΏۃR�[�h���󂯎��z��j
'   �߂�l�FBollean�i����=True/�ُ�=False�j
' ----------------------------------------------------------------------
Function FileSearch(TargetPath As String, PublishCode() As String) As Boolean
    '�ϐ��錾
    Dim Fso As New FileSystemObject
    Dim Fo As File
    Dim Ts As TextStream
    Dim Cnt As Integer
    Dim SearchResult As Integer
    Dim Ext As String
    Dim Line As String
    Dim TargetCode As String
    
    '�֐��̖߂�l��ݒ�
    FileSearch = False
    
    '�t�H���_����html�t�@�C��������
    Cnt = 0
    For Each Fo In Fso.GetFolder(TargetPath).Files
        Ext = Fso.GetExtensionName(Fo.Path)
        If (Ext = EXT_HTM Or Ext = EXT_HTML) Then
            If (FileSearch = False) Then FileSearch = True
        
            '�t�@�C�����J��
            Set Ts = Fso.OpenTextFile(Fo.Path, ForReading)
            
            '�t�@�C���̏I�[�܂�1�s���ǂݍ���
            Do Until Ts.AtEndOfStream
                Line = Ts.ReadLine
                'html�t�@�C�����烊���N�^�O����������
                SearchResult = LinkSerach(Line, TargetCode)
                
                If (SearchResult <> RESULT.NONE) Then
                    ReDim Preserve PublishCode(2, Cnt)
                    PublishCode(COL.FILE_NAME, Cnt) = Fo.Name
                    '�����N���Ɋ֘A���i���ރR�[�h�����݂��Ă����ꍇ
                    If (SearchResult = RESULT.RELATIONAL_CODE) Then
                        PublishCode(COL.RELATIONAL_CODE, Cnt) = TargetCode
                        PublishCode(COL.PRODUCT_CODE, Cnt) = ""
                    '�����N���ɏ��i�R�[�h�����݂��Ă����ꍇ
                    ElseIf (SearchResult = RESULT.PRODUCT_CODE) Then
                        PublishCode(COL.RELATIONAL_CODE, Cnt) = ""
                        PublishCode(COL.PRODUCT_CODE, Cnt) = TargetCode
                    End If
                    Cnt = Cnt + 1
                End If
            Loop
             
            ' �t�@�C�������
            Ts.Close
        End If
    Next Fo
    
    '�I�u�W�F�N�g���
    Set Ts = Nothing
    Set Fo = Nothing
    Set Fso = Nothing
End Function

' ----------------------------------------------------------------------
'   �֐����FLinkSerach�ihtml���̃n�C�p�[�����N���������j
'   ����1�F Sentence�ihtml�t�@�C������1�s�j
'   ����2�F TargetCode�ihtml�t�@�C�����̎擾�ΏۃR�[�h���󂯎��ϐ��j
'   �߂�l�FInteger�i�񋓌^RESULT�̌�����ʁj
' ----------------------------------------------------------------------
Function LinkSerach(Sentence As String, TargetCode As String) As Integer
    '�ϐ��錾
    Dim RelCodePos As Variant
    Dim ProductCodePos As Variant
    Dim StartPos As Variant
    Dim EndPos As Variant

    '�֐��̖߂�l��ݒ�
    LinkSerach = RESULT.NONE

    If InStr(Sentence, LINK_TAG) >= 1 Then
        '�֘A���i���ރR�[�h�̈ʒu���擾
        RelCodePos = InStr(Sentence, PRE_RELATIONAL_CODE)
        
        '���i�R�[�h�̈ʒu���擾
        ProductCodePos = InStr(Sentence, PRE_PRODUCT_CODE)
        
        '�ΏۃR�[�h�����݂��Ă����ꍇ�́A�R�[�h���擾
        If (RelCodePos >= 1) Then
            LinkSerach = RESULT.RELATIONAL_CODE
            StartPos = RelCodePos + Len(PRE_RELATIONAL_CODE)
            EndPos = InStr(StartPos, Sentence, AMPERSAND)
            TargetCode = Mid(Sentence, StartPos, (EndPos - StartPos))
        ElseIf (ProductCodePos >= 1) Then
            LinkSerach = RESULT.PRODUCT_CODE
            StartPos = ProductCodePos + Len(PRE_PRODUCT_CODE)
            EndPos = InStr(StartPos, Sentence, AMPERSAND)
            TargetCode = Mid(Sentence, StartPos, (EndPos - StartPos))
        End If
    End If
End Function

' ----------------------------------------------------------------------
'   �֐����FMakePublishData�i�擾�ΏۃR�[�h�����[�J���e�[�u���Ƃ��č쐬���鏈���j
'   �����F PublishCode�i�擾�ΏۃR�[�h�̔z��j
'   �߂�l�F�Ȃ�
' ----------------------------------------------------------------------
Sub MakePublishData(PublishCode() As String)
    '�ϐ��錾
    Dim TargetTable As DAO.Recordset
    Dim Cnt As Integer
    
    '�I�u�W�F�N�g�擾
    Set TargetTable = CurrentDb.OpenRecordset(TABLE_PUBLISH_CODE, dbOpenTable)
    
    '�e�[�u����������
    DoCmd.RunSQL SQL_DELETE & TABLE_PUBLISH_CODE
    
    With TargetTable
        '�e�[�u���Ƀf�[�^��ǉ�
        For Cnt = 0 To UBound(PublishCode, 2)
            .AddNew
            !No = Cnt
            !FileName = PublishCode(COL.FILE_NAME, Cnt)
            !RelationalCode = PublishCode(COL.RELATIONAL_CODE, Cnt)
            !ProductCode = PublishCode(COL.PRODUCT_CODE, Cnt)
            .Update
        Next Cnt
        '�I�u�W�F�N�g�����
        .Close
    End With
    
    '�I�u�W�F�N�g���
    Set TargetTable = Nothing
End Sub

' ----------------------------------------------------------------------
'   �֐����FGetProductData�i���i���擾�����j
'   ����1�F CurPath�i�J�����g�t�H���_�̃p�X�j
'   ����2�F CurDate�i���݂̓��t�j
'   ����3�F ConvFilePath�iCSV�t�@�C�����G�N�Z���t�@�C���ɕϊ����ꂽ��̃p�X���󂯎��ϐ��j
'   �߂�l�F�Ȃ�
' ----------------------------------------------------------------------
Sub GetProductData(CurPath As String, CurDate As String, ConvFilePath As String)
    '�ϐ��錾
    Dim ExApp As Excel.Application
    Dim TargetBook As Workbook
    Dim CopyFilePath As String
    
    '���i���t�@�C�����R�s�[
    CopyFilePath = CurPath & COPY_FILE_NAME
    If (Dir(CopyFilePath) <> "") Then Kill CopyFilePath
    FileCopy Forms!F_StartUp!TextBox10.Value, CopyFilePath
    
    '�G�N�Z���I�u�W�F�N�g���擾
    Set ExApp = CreateObject("Excel.application")
    
    '�R�s�[����CSV�t�@�C�����J��
    Set TargetBook = ExApp.Workbooks.Open(CopyFilePath)
    
    '�R�s�[����CSV�t�@�C�����G�N�Z���t�@�C���ŕۑ�
    ConvFilePath = CurPath & "\" & LINK_FILE_NAME & "_" & CurDate & EXT_EXCEL
    If (Dir(ConvFilePath) <> "") Then Kill ConvFilePath
    With TargetBook
        .SaveAs FileName:=ConvFilePath, FileFormat:=xlOpenXMLWorkbook
        .Close
    End With
    
    '�R�s�[����CSV�t�@�C�����폜
    Kill CopyFilePath
    
    '�I�u�W�F�N�g���
    Set TargetBook = Nothing
    ExApp.Quit
    Set ExApp = Nothing
End Sub

' ----------------------------------------------------------------------
'   �֐����FExportBook�i���i�����G�N�Z���t�@�C���o�͂��鏈���j
'   ����1�F CurPath�i�J�����g�t�H���_�̃p�X�j
'   ����2�F CurDate�i���݂̓��t�j
'   ����3�F ConvFilePath�iCSV�t�@�C�����G�N�Z���t�@�C���ɕϊ����ꂽ��̃p�X�j
'   ����4�F ExpFilePath�i�o�͂���G�N�Z���t�@�C���̃p�X���󂯎��ϐ��j
'   �߂�l�FBollean�i����=True/�ُ�=False�j
' ----------------------------------------------------------------------
Function ExportBook(CurPath As String, CurDate As String, ConvFilePath As String, ExpFilePath As String) As Boolean
    '�ϐ��錾
    Dim ExApp As Excel.Application
    Dim TargetBook As Workbook
    
    '�֐��̖߂�l��ݒ�
    ExportBook = True
    
    '�I�u�W�F�N�g���擾
    Set ExApp = CreateObject("Excel.application")
    Set TargetBook = ExApp.Workbooks.Open(ConvFilePath)
    
    '�G�N�Z���t�@�C���̏�����ҏW�i������j
    If (SettingFormat(TargetBook, WORD_FORM_STR, PROC_FLG.F_STR) = False) Then
        MsgBox NOT_FOUND_STR_COL, vbOKOnly, MSG_TITLE
        ExportBook = False
        Exit Function
    End If
    
    '�G�N�Z���t�@�C���̏�����ҏW�i���l�j
    If (SettingFormat(TargetBook, WORD_FORM_INT, PROC_FLG.F_INT) = False) Then
        MsgBox NOT_FOUND_INT_COL, vbOKOnly, MSG_TITLE
        ExportBook = False
        Exit Function
    End If
    
    '�G�N�Z���t�@�C���̏�����ҏW�i���t�j
    If (SettingFormat(TargetBook, WORD_FORM_DATE, PROC_FLG.F_DATE) = False) Then
        MsgBox NOT_FOUND_DATE_COL, vbOKOnly, MSG_TITLE
        ExportBook = False
        Exit Function
    End If
    
    '�t�@�C����ۑ����ĕ���
    With TargetBook
        .Sheets(TOP_SHEET).Cells(1, 1).Select
        .Close SaveChanges:=True
    End With
    
    '�G�N�Z���t�@�C���ɑ΂��郊���N�e�[�u�����쐬
    DoCmd.TransferSpreadsheet acLink, acSpreadsheetTypeExcel12Xml, LINK_FILE_NAME, ConvFilePath, True

    '�N�G�����s�i���[�J���e�[�u���쐬�j
    DoCmd.OpenQuery QUERY_C10
    
    '�����N�e�[�u���������A�G�N�Z���t�@�C�����폜
    DoCmd.DeleteObject acTable, LINK_FILE_NAME
    Kill ConvFilePath
    
    '�N�G�����G�N�Z���t�@�C���ŏo��
    ExpFilePath = CurPath & "\" & EXP_FILE_NAME & "_" & CurDate & EXT_EXCEL
    If (Dir(ExpFilePath) <> "") Then Kill ExpFilePath
    DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel12Xml, EXP_FILE_NAME, ExpFilePath, True
    DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel12Xml, EXP_DATA_NAME, ExpFilePath, True
    
    '�I�u�W�F�N�g���
    Set TargetBook = Nothing
    ExApp.Quit
    Set ExApp = Nothing
End Function

' ----------------------------------------------------------------------
'   �֐����FSettingFormat�i�����N�e�[�u���ƂȂ�G�N�Z���t�@�C���̏����ҏW�����j
'   ����1�F TargetBook�i�ҏW�Ώۂ̃G�N�Z���t�@�C���j
'   ����2�F WordList�i�J���}��؂�̃J�������j
'   ����3�F FormatTypeFlg�i�񋓌^PROC_FLG�̏�����ʃt���O�j
'   �߂�l�FBollean�i����=True/�ُ�=False�j
' ----------------------------------------------------------------------
Function SettingFormat(TargetBook As Workbook, WordList As String, FormatTypeFlg As Integer) As Boolean
    '�ϐ��錾
    Dim TargetSheet As Worksheet
    Dim AllRowRange As Range
    Dim HeaderRange As Range
    Dim TargetCell As Range
    Dim StartCell As Range
    Dim EndCell As Range
    Dim TargetWord As Variant
    
    '�֐��̖߂�l��ݒ�
    SettingFormat = True
    
    '�I�u�W�F�N�g�擾
    Set TargetSheet = TargetBook.Sheets(TOP_SHEET)
    Set AllRowRange = TargetSheet.Rows
    
    With TargetSheet
        '�w�b�_�[�͈̔͂��擾
        Set HeaderRange = .Range(.Cells(1, 1), .Cells(1, 1).End(xlToRight))

        For Each TargetWord In Split(WordList, ",")
            '�w�b�_�[�͈͂�����
            Set TargetCell = HeaderRange.Find(What:=TargetWord, LookAt:=xlWhole)
            If TargetCell Is Nothing Then
                SettingFormat = False
                Exit Function
            End If
            
            '������ݒ肷��Z�����擾
            Set StartCell = .Cells(TargetCell.Offset(1, 0).Row, TargetCell.Column)
            Set EndCell = .Cells(.Cells(AllRowRange.Count, 1).End(xlUp).Row, TargetCell.Column)

            '�����ݒ�
            With .Range(StartCell, EndCell)
                If (FormatTypeFlg = PROC_FLG.F_STR) Then
                    .NumberFormatLocal = FORM_STR
                ElseIf (FormatTypeFlg = PROC_FLG.F_INT) Then
                    .NumberFormatLocal = FORM_INT
                ElseIf (FormatTypeFlg = PROC_FLG.F_DATE) Then
                    .NumberFormatLocal = FORM_DATE
                End If
            End With
        Next TargetWord
    End With
    
    '�I�u�W�F�N�g���
    Set TargetSheet = Nothing
    Set AllRowRange = Nothing
    Set HeaderRange = Nothing
    Set TargetCell = Nothing
    Set StartCell = Nothing
    Set EndCell = Nothing
End Function

' ----------------------------------------------------------------------
'   �֐����FEditBook�i�o�̓G�N�Z���t�@�C���ҏW�����j
'   �����F ExpFilePath�i�o�͂����G�N�Z���t�@�C���̃p�X�j
'   �߂�l�F�Ȃ�
' ----------------------------------------------------------------------
Sub EditBook(ExpFilePath As String)
    '�ϐ��錾
    Dim ExApp As Excel.Application
    Dim TargetBook As Workbook
    Dim TargetSheet As Worksheet
    Dim HeaderRange As Range
    
    '�G�N�Z���I�u�W�F�N�g���擾
    Set ExApp = CreateObject("Excel.application")

    Set TargetBook = ExApp.Workbooks.Open(ExpFilePath)
    For Each TargetSheet In TargetBook.Sheets
        With TargetSheet
            .Select
            .Cells(1, 1).Select
            '�w�b�_�[�͈̔͂��擾
            Set HeaderRange = .Range(.Cells(1, 1), .Cells(1, 1).End(xlToRight))
        End With
        '�w�b�_�[�̐ݒ�i��������/�񕝎����ݒ�j
        With HeaderRange
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .AutoFilter
            .Columns.EntireColumn.AutoFit
        End With
        '�\����ʂ̐ݒ�i80%�k��/�w�b�_�[�Œ�j
        With ExApp.ActiveWindow
            .Zoom = ZOOM_POINT
            .SplitRow = HEADER_ROW
            .FreezePanes = True
        End With
    Next
    
    '�t�@�C�����㏑���ۑ����ĕ���
    With TargetBook
        .Sheets(TOP_SHEET).Select
        .Save
        .Close
    End With
    
    '�I�u�W�F�N�g���
    Set TargetBook = Nothing
    ExApp.Quit
    Set ExApp = Nothing
End Sub

' ----------------------------------------------------------------------
'   �t�@�C�����F���[�쐬
'   ������F	Excel
'   �@�\�����F  ����f�[�^�ƒ��[�e���v���[�g���i�[���ꂽ�t�H���_���w��
'				2�̃t�@�C�������ɒ��[���쐬
' ----------------------------------------------------------------------

'�萔��`
'�\�����b�Z�[�W
Const MSG_ST As String = "�������J�n���܂��B�t�H���_���w�肵�ĉ������B"
Const DIR_TITLE As String = "�Ώۃt�H���_��I��"
Const MSG_END As String = "�������I�����܂���"
Const MSG_NF As String = "��������܂���"
'�t�@�C����
Const FN_SALES As String = "01_����f�[�^.xlsx"
Const FN_TEMP As String = "09_TEMPLATE.xlsx"
Const FN_OUTPUT As String = "�T���v�����[_?????���_.xlsx"
'����
Const FML_ST As String = "=SUBTOTAL(9, ?????)"
'������
Const SHEET_TITLE As String = "�x�X�ʔ���ꗗ"
Const SUM_BR As String = "�x�X�v"
Const SUM_TEAM As String = "�`�[���v"
Const TEMP_WORD As String = "?????"
'����
Const COMMA As String = "#,##0"
Const THOU As String = "#,##0,"
Const STR_YMD As String = "YYYY�NMM��DD��"
Const INT_YMD As String = "yymmdd"
Const YM As String = "yyyy�Nmm��"
'�w�b�_�[
Const HEAD_MO As Long = 5
Const HEAD_CNT As Long = 13
Const LAST_MO As String = "�ȍ~�i13�����ڈȍ~�j"
Const THIS_MO As String = "�i�����j"
Const MO_UNIT As String = "�����ځj"
'���̑����l
Const TITLE_POS As String = "A1"
Const COPY_COEF As Long = 3
Const PASTE_COEF As Long = 4
Const ELEM_CNT As Long = 2
Const ITEM_CNT As Long = 14
Const ITEM_END As Long = 62
Const DATA_START As Long = 5
Const GT As Long = 7
Const BEGIN As Long = 9
Const GREY As Long = 15

'�񋓌^��`
'�o�̓t�@�C���̗�
Enum OUTPUT
    CODE = 2
    BRANCH = 3
    TEAM = 4
    BOOK = 6
End Enum

'�����t���O
Enum FLG
    BR_ST = 1
    TEAM_ST = 2
End Enum

'���[�U�[��`�^�錾
'��ƃt�@�C��
Type Booklist
    Data As Workbook
    Main As Workbook
End Type

'���t
Type Datelist
    BaseDate As Date
    BaseDateStr As String
End Type

' ----------------------------------------------------------------------
'   �֐����FMainProc�i�又���j
'   �����F  �Ȃ�
'   �߂�l�F�Ȃ�
' ----------------------------------------------------------------------
Sub MainProc()
    '�ϐ��錾
    Dim TargetBook As Booklist
    Dim FolderPath As String
    Dim CopyFile As String
    Dim ToCopyPath As String
    Dim TargetRng As Range
    
    '���s�m�F
    If MsgBox(MSG_ST, vbOKCancel + vbInformation) = vbCancel Then Exit Sub
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = DIR_TITLE                      '�^�C�g���̎w��
        .InitialFileName = ThisWorkbook.Path    '�����\���t�H���_�̎w��
        If .Show = True Then                    '�_�C�A���O��\�����Ė߂�l�𔻒�
            FolderPath = .SelectedItems(1)      '�t�H���_�̃p�X���擾
        Else
            Exit Sub
        End If
    End With
    
    '��������
    If InitProc(TargetBook, FolderPath) = False Then Exit Sub
    
    '�e���v���[�g�t�@�C����ʖ��ŕۑ�
    CopyFile = Replace(FN_OUTPUT, TEMP_WORD, Format(Date, INT_YMD))
    ToCopyPath = ThisWorkbook.Path & "\" & CopyFile
    FileCopy FolderPath & "\" & FN_TEMP, ToCopyPath
    
    '�ʖ��ۑ������t�@�C�����J��
    Set TargetBook.Main = Workbooks.Open(ToCopyPath)
    
    '����f�[�^�̃R�s�[�Ɠ\��t��
    Call GetSalesData(TargetBook)
    
    '���v�s�쐬
    With ActiveSheet
        Set TargetRng = .Range(.Cells(BEGIN, OUTPUT.CODE), .Cells(Rows.Count, OUTPUT.CODE).End(xlUp))
    End With
    Call MakeSubTotal(TargetRng, FLG.BR_ST)
    
    '�����ݒ�
    Call EditFormat(TargetBook.Main.Sheets(1))
    
    '�t�@�C�������
    With TargetBook.Main
        .Sheets(1).Activate
        .Sheets(1).Range(TITLE_POS).Select
        .Close SaveChanges:=True
    End With
    Set TargetBook.Main = Nothing
    
    '�I������
    Call EndProc
End Sub

' ----------------------------------------------------------------------
'   �֐����FInitProc�i���������j
'   ����1�F TargetBook�i�G�N�Z���t�@�C�����X�g�j
'   ����2�F FolderPath�i���[�U�[���w�肵���t�H���_�̃p�X�j
'   �߂�l�FBollean�i����=True/�ُ�=False�j
' ----------------------------------------------------------------------
Function InitProc(TargetBook As Booklist, FolderPath As String) As Boolean
    
    With Application
        '��ʍX�V��~
        .ScreenUpdating = False
        '�m�F���b�Z�[�W��\��
        .DisplayAlerts = False
    End With
    
    '�J�����g�f�B���N�g���ύX
    ChDir ThisWorkbook.Path

    '01_����f�[�^.xlsx�t�@�C�����J��
    If (Dir(FolderPath & "\" & FN_SALES) = "") Then
        MsgBox FN_SALES & MSG_NF, vbCritical
        GoTo ErrProc
    End If
    Set TargetBook.Data = Workbooks.Open(FolderPath & "\" & FN_SALES)
       
    '09_TEMPLATE.xlsx�����݂��邱�Ƃ��m�F
    If (Dir(FolderPath & "\" & FN_TEMP) = "") Then
        MsgBox FN_TEMP & MSG_NF, vbCritical
        TargetBook.Data.Close
        GoTo ErrProc
    End If
    
    '����I��
    InitProc = True
    Exit Function
    
ErrProc:  '�G���[����
    Call EndProc
    InitProc = False
    Exit Function
End Function

' ----------------------------------------------------------------------
'   �֐����FGetSalesData�i����f�[�^�擾�����j
'   �����F  TargetBook�i�G�N�Z���t�@�C�����X�g�j
'   �߂�l�F�Ȃ�
' ----------------------------------------------------------------------
Sub GetSalesData(TargetBook As Booklist)
    '�ϐ��錾
    Dim TargetSheet As Worksheet
    Dim CopyRng As Range
    Dim CopyStPos As Long
    Dim PasteStPos As Long
    Dim MainCnt As Long
    Dim SubCnt As Long
    Dim ColChar As String
    Dim TargetCol As Long
    Dim StRngStr As String
    
    '�w�b�_�[��ݒ�
    Set TargetSheet = TargetBook.Main.Sheets(1)
    Call EditHeader(TargetSheet)
    
    '����f�[�^�͈̔͂��擾
    With TargetBook.Data.Sheets(1).UsedRange
        Set CopyRng = .Offset(1).Resize(.Rows.Count - 1)
    End With
    
    '�o�̓t�@�C���Ɍ��o����]�L
    CopyRng.Resize(, 4).Copy
    TargetSheet.Cells(BEGIN, OUTPUT.CODE).PasteSpecial Paste:=xlPasteValues
    '�r���ݒ�
    Call EditLine(Selection)
    Selection.End(xlDown).Offset(1).Resize(, 4).Borders(xlEdgeTop).LineStyle = xlContinuous
    
    '�o�̓t�@�C���ɔ���f�[�^��]�L
    For MainCnt = 0 To ITEM_CNT
        CopyStPos = DATA_START + (COPY_COEF * MainCnt) - 1
        PasteStPos = OUTPUT.BOOK + (PASTE_COEF * MainCnt)
        
        '�R�s�[���Đ��l�`���œ\��t��
        CopyRng.Offset(, CopyStPos).Resize(, 3).Copy
        TargetSheet.Cells(BEGIN, PasteStPos).PasteSpecial Paste:=xlPasteValues
        
        '�r���ݒ�
        Call EditLine(Selection)
        Selection.End(xlDown).Offset(1).Resize(, 3).Borders(xlEdgeTop).LineStyle = xlContinuous
        '���v�s�ɐ��������
        For SubCnt = 0 To ELEM_CNT
            TargetCol = PasteStPos + SubCnt
            ColChar = Replace(TargetSheet.Cells(1, TargetCol).Address(True, False), "$1", "")
            With TargetSheet.Cells(BEGIN, TargetCol)
                StRngStr = ColChar & .Row & ":" & ColChar & .End(xlDown).Row
            End With
            TargetSheet.Cells(GT, TargetCol).Formula = Replace(FML_ST, TEMP_WORD, StRngStr)
        Next SubCnt
    Next MainCnt
    
    '�I�u�W�F�N�g���
    Set TargetSheet = Nothing
    Set CopyRng = Nothing
    
    '�t�@�C�������
    TargetBook.Data.Close
    Set TargetBook.Data = Nothing
End Sub

' ----------------------------------------------------------------------
'   �֐����FEditHeader�i�w�b�_�[�ݒ菈���j
'   �����F  TargetSheet�i�ΏۃG�N�Z���V�[�g�j
'   �߂�l�F�Ȃ�
' ----------------------------------------------------------------------
Sub EditHeader(TargetSheet As Worksheet)
    '�ϐ��錾
    Dim Cnt As Long
    Dim TargetCol As Long
    Dim TargetDate As Datelist
    Dim HeaderDate As String
    Dim HeadStr As String
    
    '���t���擾
    With TargetDate
        .BaseDate = Date
        .BaseDateStr = Format(.BaseDate, STR_YMD)
    End With
    
    '�^�C�g����ݒ�
    With TargetSheet.Range(TITLE_POS)
        .Value = Replace(.Value, TEMP_WORD, SHEET_TITLE & "_" & TargetDate.BaseDateStr)
    End With
    
    '�w�b�_�[��ݒ�
    For Cnt = 0 To HEAD_CNT
        TargetCol = ITEM_END - (PASTE_COEF * Cnt)
        If (Cnt = 0) Then
            HeaderDate = Format(DateAdd("m", HEAD_CNT, TargetDate.BaseDate), YM)
            HeadStr = HeaderDate & LAST_MO
        ElseIf (Cnt = HEAD_CNT) Then
            HeaderDate = Format(TargetDate.BaseDate, YM)
            HeadStr = HeaderDate & THIS_MO
        Else
            HeaderDate = Format(DateAdd("m", HEAD_CNT - Cnt, TargetDate.BaseDate), YM)
            HeadStr = HeaderDate & "�i" & HEAD_CNT - Cnt & MO_UNIT
        End If
        TargetSheet.Cells(HEAD_MO, TargetCol).Value = HeadStr
    Next Cnt
End Sub

' ----------------------------------------------------------------------
'   �֐����FEditLine�i�r���ݒ菈���j
'   �����F  TargetRng�i�Ώ۔͈́j
'   �߂�l�F�Ȃ�
' ----------------------------------------------------------------------
Sub EditLine(TargetRng As Range)
    With TargetRng
        .Borders(xlEdgeLeft).LineStyle = xlContinuous
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeRight).LineStyle = xlContinuous
        .Borders(xlInsideVertical).LineStyle = xlContinuous
        .Borders(xlInsideHorizontal).LineStyle = xlContinuous
    End With
End Sub

' ----------------------------------------------------------------------
'   �֐����FMakeSubTotal�i���v�s�쐬�����j
'   ����1�F TargetRng�i�Ώ۔͈́j
'   ����2�F ProcFlg�i�񋓌^FLG�̏����t���O�j
'   �߂�l�F�Ȃ�
' ----------------------------------------------------------------------
Sub MakeSubTotal(TargetRng As Range, ProcFlg As Long)
    '�ϐ��錾
    Dim TargetCell As Range
    Dim StCell As Range
    Dim EndCell As Range
    Dim BrRng As Range
    Dim TeamRng As Range
    Dim LastValue As Variant
    
    For Each TargetCell In TargetRng
        If (LastValue <> TargetCell.Value) Then
            '�l������
            With TargetRng
                Set EndCell = .Find(What:=TargetCell.Value, SearchDirection:=xlPrevious, LookAt:=xlWhole)
                Set StCell = .Find(What:=TargetCell.Value, After:=EndCell, SearchDirection:=xlNext, LookAt:=xlWhole)
            End With
            
            '�����Ə����̐ݒ�
            Call EditFormulaAndFormat(StCell, EndCell, ProcFlg)
            
            If (ProcFlg = FLG.BR_ST) Then
                '�Ώۃ`�[���͈͂��擾
                With ActiveSheet
                    Set TeamRng = .Range(.Cells(StCell.Row, OUTPUT.TEAM), .Cells(EndCell.Row, OUTPUT.TEAM))
                End With
                '�ċA����
                Call MakeSubTotal(TeamRng, FLG.TEAM_ST)
                
                '�Ώێx�X�͈͂��擾
                With ActiveSheet
                    Set BrRng = .Range(.Cells(StCell.Row, OUTPUT.CODE), .Cells(EndCell.Row, OUTPUT.BRANCH))
                End With
                '������ݒ�
                With BrRng
                    .Offset(-1).Resize(.Rows.Count + 1).Rows.Group
                    .Offset(-1).Resize(.Rows.Count + 1).Font.ColorIndex = GREY
                    .Offset(-2).Resize(.Rows.Count + 2).Borders(xlInsideHorizontal).LineStyle = xlLineStyleNone
                End With
            End If
            
            '�l��ۑ�
            LastValue = TargetCell.Value
        End If
    Next TargetCell
    
    '�I�u�W�F�N�g���
    Set TargetRng = Nothing
    Set TargetCell = Nothing
    Set StCell = Nothing
    Set EndCell = Nothing
    Set BrRng = Nothing
    Set TeamRng = Nothing
End Sub

' ----------------------------------------------------------------------
'   �֐����FEditFormulaAndFormat�i�����Ə����̐ݒ菈���j
'   ����1�F RngSt�i�͈͎n�_�j
'   ����2�F RngEnd�i�͈͏I�_�j
'   ����3�F ProcFlg�i�񋓌^FLG�̏����t���O�j
'   �߂�l�F�Ȃ�
' ----------------------------------------------------------------------
Sub EditFormulaAndFormat(RngSt As Range, RngEnd As Range, ProcFlg As Long)
    '�ϐ��錾
    Dim MainCnt As Long
    Dim SubCnt As Long
    Dim TargetCol As Long
    Dim ColChar As String
    Dim RowFml As Long
    Dim RowSt As Long
    Dim RowEnd As Long
    Dim BaseRng As Range
    Dim SubtRng As String
    
    '�s��}��
    RngSt.EntireRow.Insert
    '�n�_�ƏI�_�̍s�ԍ�
    RowSt = RngSt.Row
    RowEnd = RngEnd.Row
    '�}���s�̍s�ԍ��i�������͍s�j
    RowFml = RowSt - 1
    
    '��͈͂��擾
    With ActiveSheet
        Set BaseRng = .Range(.Cells(RowSt, OUTPUT.CODE), .Cells(RowEnd, OUTPUT.CODE))
    End With
    
    With BaseRng
        If (ProcFlg = FLG.BR_ST) Then
            '�}���s�ɃR�s�[�����l��\��t��
            .Resize(1, 2).Copy
            .Offset(-1).Resize(1).PasteSpecial Paste:=xlPasteValues
            '�����̐ݒ��������
            .Offset(-1).Resize(1, 4).ClearFormats
            '�����𑾎��ɐݒ�
            With .Offset(-1, 2).Resize(1)
                .Value = SUM_BR
                .Font.Bold = True
            End With
            '�r���ݒ�
            Call EditLine(.Offset(-1).Resize(1, 4))
            .Offset(-1, 2).Resize(1, 2).Borders(xlInsideVertical).LineStyle = xlLineStyleNone
        ElseIf (ProcFlg = FLG.TEAM_ST) Then
            '�}���s�ɃR�s�[�����l��\��t��
            .Resize(1, 3).Copy
            .Offset(-1).Resize(1).PasteSpecial Paste:=xlPasteValues
            '�����̐ݒ��������
            .Offset(-1, 2).Resize(1, 2).ClearFormats
             '�����𑾎��ɐݒ�
            With .Offset(-1, 3).Resize(1)
                .Value = SUM_TEAM
                .Font.Bold = True
            End With
            '�r���ݒ�
            Call EditLine(.Offset(-1, 2).Resize(, 2))
            .Offset(-1, 2).Resize(.Rows.Count + 1).Borders(xlInsideHorizontal).LineStyle = xlLineStyleNone
            With .Offset(, 3).Resize(, 4)
                .Borders(xlInsideHorizontal).LineStyle = xlDash
                .Borders(xlInsideHorizontal).Weight = xlThin
                '�O���[�v��
                .Rows.Group
            End With
            '���v�s�ȊO�̕����F��ݒ�
            .Offset(, 2).Font.ColorIndex = GREY
        End If
    End With
        
    '�}���s�̕���ݒ�
    With ActiveSheet.Rows(RowFml)
        .EntireRow.AutoFit
        .ClearOutline
    End With

    For MainCnt = 0 To ITEM_CNT
        '�Ώۂ̗�ԍ��擾
        TargetCol = OUTPUT.BOOK + (PASTE_COEF * MainCnt)
        
        '���������
        For SubCnt = 0 To ELEM_CNT
            With ActiveSheet.Cells(RowFml, TargetCol + SubCnt)
                ColChar = Replace(.Address(True, False), "$" & RowFml, "")
                SubtRng = ColChar & RowSt & ":" & ColChar & RowEnd
                .Formula = Replace(FML_ST, TEMP_WORD, SubtRng)
            End With
        Next SubCnt
        
        '��͈͂��擾
        With ActiveSheet
            Set BaseRng = .Range(.Cells(RowSt, TargetCol), .Cells(RowEnd, TargetCol))
        End With
        
        '�����̐ݒ�
        With BaseRng
            If (ProcFlg = FLG.BR_ST) Then
                '�����̐ݒ��������
                .Offset(-1).Resize(1, 3).ClearFormats
                '�r���ݒ�
                Call EditLine(.Offset(-1).Resize(1, 3))
            ElseIf (ProcFlg = FLG.TEAM_ST) Then
                '�r���ݒ�
                With .Resize(, 3)
                    .Borders(xlInsideHorizontal).LineStyle = xlDash
                    .Borders(xlInsideHorizontal).Weight = xlThin
                End With
            End If
            '�����𑾎��ɐݒ�
            .Offset(-1).Resize(1, 3).Font.Bold = True
        End With
    Next MainCnt
    
    '�I�u�W�F�N�g���
    Set BaseRng = Nothing
End Sub

' ----------------------------------------------------------------------
'   �֐����FEditFormat�i�����ݒ菈���j
'   �����F  TargetSheet�i�ΏۃG�N�Z���V�[�g�j
'   �߂�l�F�Ȃ�
' ----------------------------------------------------------------------
Sub EditFormat(TargetSheet As Worksheet)
    '�ϐ��錾
    Dim Cnt As Long
    Dim BaseRng As Range
    Dim EditStPos As Long
    
    '��͈͂��擾
    With TargetSheet
        Set BaseRng = .Range(.Cells(BEGIN, OUTPUT.CODE), .Cells(Rows.Count, OUTPUT.CODE).End(xlUp))
    End With
        
    For Cnt = 0 To ITEM_CNT
        EditStPos = PASTE_COEF * (Cnt + 1)
        '�����ݒ�
        With BaseRng.Offset(, EditStPos)
            .Resize(, 2).NumberFormatLocal = COMMA
            .Offset(, 2).NumberFormatLocal = THOU
        End With
        '�����t�������ݒ�i�l��0�̏ꍇ�A�����F���D�F�j
        With BaseRng.Offset(, EditStPos).Resize(, 3)
            .FormatConditions.Add Type:=xlCellValue, Operator:=xlEqual, Formula1:="=0"
            .FormatConditions(1).Font.ColorIndex = GREY
        End With
    Next Cnt
    '�l�𒆉�����
    BaseRng.Resize(, 2).HorizontalAlignment = xlCenter
    '�O���[�v�s��܂���
    TargetSheet.Outline.ShowLevels RowLevels:=2
    
    '�I�u�W�F�N�g���
    Set BaseRng = Nothing
End Sub

' ----------------------------------------------------------------------
'   �֐����FEndProc�i�}�N���I�������j
'   �����F  �Ȃ�
'   �߂�l�F�Ȃ�
' ----------------------------------------------------------------------
Sub EndProc()
    With Application
        '��ʍX�V�ĊJ
        .ScreenUpdating = True
    
        '�m�F���b�Z�[�W�\��
        .DisplayAlerts = True
    End With
    
    '�I�����b�Z�[�W
    MsgBox MSG_END, vbInformation
End Sub

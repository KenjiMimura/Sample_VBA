' ----------------------------------------------------------------------
'   ファイル名：ファイル名一括置換
'   動作環境：	Excel
'   機能説明：  指定したフォルダ内の全ファイルのファイル名を指定した文言で置換
' ----------------------------------------------------------------------

'定数定義
'メッセージ関連
Const DIR_TITLE As String = "対象フォルダを選択"
Const MSG_SELECT As String = "ファイル名を置換したいファイルが格納されたフォルダを指定して下さい。" & vbCrLf & _
								"※　指定したフォルダ内の全ファイルが処理対象となります。"
Const MSG_FIND As String = "置換の対象となる文字列を入力して下さい。"
Const MSG_REPLACE As String = "置換後の文字列を入力して下さい。"
Const MSG_ERR As String = "ファイル名として使用出来ない文字が含まれています。" & vbCrLf & "もう一度、入力して下さい。"
Const MSG_CONF As String = "この内容でファイル名の置換を実行しますが、よろしいですか？"
Const MSG_SUCCESS As String = "ファイル名変更成功"
Const MSG_FAIL As String = "ファイル名変更失敗"
Const MSG_SUSPEND As String = "処理を終了します。"
Const MSG_END As String = "処理が完了しました。"
'実行確認画面関連
Const DLG_TITLE As String = "確認"
Const EXEC_HISTORY As String = "実行履歴"
Const TARGET_FOLDER = "対象フォルダ　：　"
Const FIND_WORD = "置換前文字列　：　"
Const REPLACE_WORD = "置換後文字列　：　"
Const RESULT_HEAD As String = "実行結果"
Const ERROR_HEAD As String = "エラー内容"
Const PATH_HEAD As String = "ファイル格納場所"
Const BF_FILE_HEAD As String = "変更前ファイル名"
Const AF_FILE_HEAD As String = "変更後ファイル名"
'その他定数
Const HEADER_ROW As Integer = 1
Const FIRST_ROW As Integer = 2
Const RESULT_COL As String = "A"
Const AF_FILE_COL As String = "E"

'列挙型定義
'ヘッダー関連
Enum HEADER
    EXEC_RESULT = 1
    EXEC_ERROR
    FILE_PATH
    BF_FILE
    AF_FILE
End Enum

' ----------------------------------------------------------------------
'   関数名：FileNameReplacement（主処理）
'   引数：  なし
'   戻り値：なし
' ----------------------------------------------------------------------
Sub FileNameReplacement()
    '変数宣言
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
        '画面更新停止
        .ScreenUpdating = False
        '確認メッセージ非表示
        .DisplayAlerts = False
    End With

    'カレントディレクトリ変更
    ChDir ThisWorkbook.Path

    '対象のフォルダを指定
    MsgBox MSG_SELECT, vbInformation
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = DIR_TITLE                      'タイトルの指定
        .InitialFileName = ThisWorkbook.Path    '初期表示フォルダの指定
        If .Show = True Then                    'ダイアログを表示して戻り値を判定
            FolderPath = .SelectedItems(1)      'フォルダのパスを取得
        Else
            Exit Sub
        End If
    End With
    
    '検索文字列入力
    If UserInput(FindStr, MSG_FIND) = False Then Exit Sub
    
    '置換時列入力
    If UserInput(ReplaceStr, MSG_REPLACE) = False Then Exit Sub
     
    '処理の実行を確認
    Result = MsgBox( _
                    TARGET_FOLDER & FolderPath & vbCrLf & vbCrLf & _
                    FIND_WORD & FindStr & vbCrLf & _
                    REPLACE_WORD & ReplaceStr & vbCrLf & vbCrLf & _
                    MSG_CONF, vbYesNo + vbQuestion, DLG_TITLE)
    '処理キャンセル時
    If Result = vbNo Then
        MsgBox MSG_SUSPEND
        Exit Sub
    End If

    '実行履歴シートが存在している場合は削除
    For SheetCnt = 1 To Worksheets.Count
        If Worksheets(SheetCnt).Name = EXEC_HISTORY Then
            Worksheets(EXEC_HISTORY).Delete
            Exit For
        End If
    Next SheetCnt
    
    '実行履歴シートを作成
    Set NewSheet = ThisWorkbook.Worksheets.Add(after:=Worksheets(Worksheets.Count))
    NewSheet.Name = EXEC_HISTORY
    
    '実行履歴シートのヘッダー行を作成
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
        '置換後のファイル名を生成
        NewFileName = Replace(File.Name, FindStr, ReplaceStr)

        If File.Name <> NewFileName Then
            On Error Resume Next
            
            With NewSheet
                .Cells(ExecCnt, HEADER.FILE_PATH).Value = File.ParentFolder
                .Cells(ExecCnt, HEADER.BF_FILE).Value = File.Name
                .Cells(ExecCnt, HEADER.AF_FILE).Value = NewFileName
                
                'ファイル名変更
                Name File.Path As File.ParentFolder & "\" & NewFileName
                
                If Err.Number <> 0 Then
                    'ファイル名の変更が出来なかった場合の処理
                    .Cells(ExecCnt, HEADER.EXEC_RESULT).Value = MSG_FAIL
                    .Cells(ExecCnt, HEADER.EXEC_ERROR).Value = Err.Description
                    .Range(RESULT_COL & ExecCnt & ":" & AF_FILE_COL & ExecCnt).Font.Color = RGB(255, 0, 0) '文字色:赤
                    Err.Clear
                Else
                    .Cells(ExecCnt, HEADER.EXEC_RESULT).Value = MSG_SUCCESS
                End If
            End With
            ExecCnt = ExecCnt + 1
        End If
    Next
    
    'オブジェクト解放
    Set Fso = Nothing
    Set File = Nothing
    
    '列を自動調節
    NewSheet.Columns(RESULT_COL & ":" & AF_FILE_COL).AutoFit

    With Application
        '画面更新停止
        .ScreenUpdating = True
        '確認メッセージ非表示
        .DisplayAlerts = True
    End With

    '終了メッセージを表示
    MsgBox MSG_END, vbInformation
End Sub

' ----------------------------------------------------------------------
'   関数名：UserInput（ユーザー入力処理）
'   引数1： InputString（ユーザーの入力を受け取る変数）
'   引数2： OutputMessage（テキストボックスに表示される文言）
'   戻り値：Bollean（正常=True/異常=False）
' ----------------------------------------------------------------------
Function UserInput(InputString As Variant, OutputMessage As String) As Boolean
    '変数宣言
    Dim CheckFlg As Boolean

    '関数の戻り値を設定
    UserInput = True
    
    'ユーザー入力
    Do
        CheckFlg = True
        InputString = InputBox(OutputMessage)
        If StrPtr(InputString) = 0 Then
            MsgBox MSG_SUSPEND
            UserInput = False
            Exit Function
        End If
        
        '使用禁止文字の確認
        If InStr(InputString, "\") > 0 Then CheckFlg = False
        If InStr(InputString, "/") > 0 Then CheckFlg = False
        If InStr(InputString, ":") > 0 Then CheckFlg = False
        If InStr(InputString, "*") > 0 Then CheckFlg = False
        If InStr(InputString, "?") > 0 Then CheckFlg = False
        If InStr(InputString, """") > 0 Then CheckFlg = False
        If InStr(InputString, "<") > 0 Then CheckFlg = False
        If InStr(InputString, ">") > 0 Then CheckFlg = False
        If InStr(InputString, "|") > 0 Then CheckFlg = False
        
        '使用禁止文字が入力された場合の処理
        If CheckFlg = False Then
            MsgBox MSG_ERR
        End If
    Loop While CheckFlg = False
End Function

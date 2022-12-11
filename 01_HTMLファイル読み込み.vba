' ----------------------------------------------------------------------
'   ファイル名：HTMLファイル読み込み
'   動作環境：	Access
'   機能説明：  指定したフォルダ内の全HTMLファイルから商品関連のコードを取得
'				取得した商品関連のコードと商品情報のCSVファイルをACCESSに取り込み
'				HTMLファイルに記載されていた商品情報をエクセルファイルで出力
' ----------------------------------------------------------------------

'定数定義
'ダイアログ関連
Const DLG_TITLE_FILE As String = "ファイルを選択"
Const DLG_TITLE_FOLDER As String = "フォルダを選択"
Const FILTER_DESC As String = "CSVファイル"
Const FILTER_EXT As String = "*.csv"
'検索処理関連
Const EXT_HTM As String = "htm"
Const EXT_HTML As String = "html"
Const LINK_TAG As String = "href"
Const AMPERSAND As String = "&"
Const PRE_RELATIONAL_CODE As String = "&relational_code="
Const PRE_PRODUCT_CODE As String = "&product_code="
Const WORD_FORM_STR As String = "RELATIONAL_CODE,RELATIONAL_TYPE,PRODUCT_CODE,PRODUCT_NAME,BRANCH_NAME,TEAM_NAME,INFLOW_ROUTE"
Const WORD_FORM_INT As String = "QUANTITY,UNIT_PRICE,SALES,INFLOW_ROUTE_NO"
Const WORD_FORM_DATE As String = "RELEASE_DATE,BOOKING_DATE"
'データベース関連
Const TABLE_PUBLISH_CODE As String = "PublishCode"
Const SQL_DELETE As String = "DELETE * FROM "
Const QUERY_C10 As String = "C10_ProductData"
'ファイル関連
Const COPY_FILE_NAME = "\Target.csv"
Const EXT_EXCEL As String = ".xlsx"
Const LINK_FILE_NAME As String = "DATA"
Const EXP_FILE_NAME As String = "掲載商品情報"
Const EXP_DATA_NAME As String = "掲載商品一覧"
'メッセージ関連
Const MSG_TITLE = "GetProductData"
Const MSG_END As String = "処理が完了しました"
'エラー関連
Const NOT_FOUND_HTML As String = "htmlファイルが見つかりませんでした。"
Const NOT_FOUND_STR_COL As String = "商品情報ファイルから対象の列が見つかりませんでした。（文字列）"
Const NOT_FOUND_INT_COL As String = "商品情報ファイルから対象の列が見つかりませんでした。（数値）"
Const NOT_FOUND_DATE_COL As String = "商品情報ファイルから対象の列が見つかりませんでした。（日付）"
'書式関連
Const FORM_DATE8 As String = "yyyymmdd"
Const FORM_STR As String = "@"
Const FORM_INT As String = "0_ "
Const FORM_DATE As String = "yyyy/mm/dd"
'数値関連
Const ZOOM_POINT As Integer = 80
Const HEADER_ROW As Integer = 1
Const TOP_SHEET As Integer = 1

'列挙型定義
'リンク検索種別
Enum RESULT
    NONE = 0
    RELATIONAL_CODE
    PRODUCT_CODE
End Enum

'PublishCodeテーブルのカラム
Enum COL
    FILE_NAME = 0
    RELATIONAL_CODE
    PRODUCT_CODE
End Enum

'書式種別フラグ
Enum PROC_FLG
    F_STR = 1
    F_INT
    F_DATE
End Enum

' ----------------------------------------------------------------------
'   関数名：Button10_Click（ファイル選択ボタン押下時処理）
'   引数：  なし
'   戻り値：なし
' ----------------------------------------------------------------------
Private Sub Button10_Click()
    'ファイル選択ダイアログボックスを表示
    With Application.FileDialog(msoFileDialogFilePicker)
        'タイトルを設定
        .Title = DLG_TITLE_FILE
        
        'ファイルの種類を設定
        .Filters.Clear
        .Filters.Add FILTER_DESC, FILTER_EXT
        .FilterIndex = 1
        
        '複数ファイル選択を許可しない
        .AllowMultiSelect = False
        
        '初期パスを設定
        If (Forms!F_StartUp!TextBox10.Value <> "") Then
            .InitialFileName = Forms!F_StartUp!TextBox10.Value
        Else
            .InitialFileName = CurrentProject.Path
        End If
        
        If (.Show <> 0) Then
            '選択されたファイルのパスをテキストボックスに設定
            Forms!F_StartUp!TextBox10.Value = Trim(.SelectedItems.Item(1))
        Else
            Exit Sub
        End If
    End With
End Sub

' ----------------------------------------------------------------------
'   関数名：Button20_Click（処理開始ボタン押下時処理）
'   引数：  なし
'   戻り値：なし
' ----------------------------------------------------------------------
Private Sub Button20_Click()
    '変数宣言
    Dim CurPath As String
    Dim CurDate As String
    Dim TargetPath As String
    Dim ConvFilePath As String
    Dim ExpFilePath As String
    ReDim PublishCode(2, 0) As String
    
    'システムメッセージ非表示
    DoCmd.SetWarnings False
    
    'カレントフォルダーのパスを取得
    CurPath = CurrentProject.Path
    
    '現在の日付をyyyymmdd形式で取得
    CurDate = FORMAT(Date, FORM_DATE8)
    
    'フォルダー選択ダイアログボックスを表示
    With Application.FileDialog(msoFileDialogFolderPicker)
        'タイトルを設定
        .Title = DLG_TITLE_FOLDER
        
        '初期パスを設定
        .InitialFileName = CurPath
        
        If (.Show <> 0) Then
            '選択されたフォルダのパスを取得
            TargetPath = Trim(.SelectedItems.Item(1))
        Else
            Exit Sub
        End If
    End With
        
    'htmlファイルを検索
    If (FileSearch(TargetPath, PublishCode) = False) Then
        MsgBox NOT_FOUND_HTML, vbOKOnly, MSG_TITLE
        Exit Sub
    End If
    
    'html内の対象コードをローカルテーブルとして作成
    MakePublishData PublishCode
    
    '商品の情報を取得
    GetProductData CurPath, CurDate, ConvFilePath
    
    'エクセルファイルに出力
    If (ExportBook(CurPath, CurDate, ConvFilePath, ExpFilePath) = False) Then Exit Sub
    
    '出力したエクセルファイルを編集
    EditBook ExpFilePath
    
    'システムメッセージ表示
    DoCmd.SetWarnings True
    
    '処理終了メッセージ
    MsgBox MSG_END, vbOKOnly, MSG_TITLE
End Sub

' ----------------------------------------------------------------------
'   関数名：FileSearch（ファイル検索処理）
'   引数1： TargetPath（ユーザーが選択したフォルダーのパス）
'   引数2： PublishCode（htmlファイル内の取得対象コードを受け取る配列）
'   戻り値：Bollean（正常=True/異常=False）
' ----------------------------------------------------------------------
Function FileSearch(TargetPath As String, PublishCode() As String) As Boolean
    '変数宣言
    Dim Fso As New FileSystemObject
    Dim Fo As File
    Dim Ts As TextStream
    Dim Cnt As Integer
    Dim SearchResult As Integer
    Dim Ext As String
    Dim Line As String
    Dim TargetCode As String
    
    '関数の戻り値を設定
    FileSearch = False
    
    'フォルダ内のhtmlファイルを検索
    Cnt = 0
    For Each Fo In Fso.GetFolder(TargetPath).Files
        Ext = Fso.GetExtensionName(Fo.Path)
        If (Ext = EXT_HTM Or Ext = EXT_HTML) Then
            If (FileSearch = False) Then FileSearch = True
        
            'ファイルを開く
            Set Ts = Fso.OpenTextFile(Fo.Path, ForReading)
            
            'ファイルの終端まで1行ずつ読み込む
            Do Until Ts.AtEndOfStream
                Line = Ts.ReadLine
                'htmlファイルからリンクタグを検索する
                SearchResult = LinkSerach(Line, TargetCode)
                
                If (SearchResult <> RESULT.NONE) Then
                    ReDim Preserve PublishCode(2, Cnt)
                    PublishCode(COL.FILE_NAME, Cnt) = Fo.Name
                    'リンク内に関連商品分類コードが存在していた場合
                    If (SearchResult = RESULT.RELATIONAL_CODE) Then
                        PublishCode(COL.RELATIONAL_CODE, Cnt) = TargetCode
                        PublishCode(COL.PRODUCT_CODE, Cnt) = ""
                    'リンク内に商品コードが存在していた場合
                    ElseIf (SearchResult = RESULT.PRODUCT_CODE) Then
                        PublishCode(COL.RELATIONAL_CODE, Cnt) = ""
                        PublishCode(COL.PRODUCT_CODE, Cnt) = TargetCode
                    End If
                    Cnt = Cnt + 1
                End If
            Loop
             
            ' ファイルを閉じる
            Ts.Close
        End If
    Next Fo
    
    'オブジェクト解放
    Set Ts = Nothing
    Set Fo = Nothing
    Set Fso = Nothing
End Function

' ----------------------------------------------------------------------
'   関数名：LinkSerach（html内のハイパーリンク検索処理）
'   引数1： Sentence（htmlファイル内の1行）
'   引数2： TargetCode（htmlファイル内の取得対象コードを受け取る変数）
'   戻り値：Integer（列挙型RESULTの検索種別）
' ----------------------------------------------------------------------
Function LinkSerach(Sentence As String, TargetCode As String) As Integer
    '変数宣言
    Dim RelCodePos As Variant
    Dim ProductCodePos As Variant
    Dim StartPos As Variant
    Dim EndPos As Variant

    '関数の戻り値を設定
    LinkSerach = RESULT.NONE

    If InStr(Sentence, LINK_TAG) >= 1 Then
        '関連商品分類コードの位置を取得
        RelCodePos = InStr(Sentence, PRE_RELATIONAL_CODE)
        
        '商品コードの位置を取得
        ProductCodePos = InStr(Sentence, PRE_PRODUCT_CODE)
        
        '対象コードが存在していた場合は、コードを取得
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
'   関数名：MakePublishData（取得対象コードをローカルテーブルとして作成する処理）
'   引数： PublishCode（取得対象コードの配列）
'   戻り値：なし
' ----------------------------------------------------------------------
Sub MakePublishData(PublishCode() As String)
    '変数宣言
    Dim TargetTable As DAO.Recordset
    Dim Cnt As Integer
    
    'オブジェクト取得
    Set TargetTable = CurrentDb.OpenRecordset(TABLE_PUBLISH_CODE, dbOpenTable)
    
    'テーブルを初期化
    DoCmd.RunSQL SQL_DELETE & TABLE_PUBLISH_CODE
    
    With TargetTable
        'テーブルにデータを追加
        For Cnt = 0 To UBound(PublishCode, 2)
            .AddNew
            !No = Cnt
            !FileName = PublishCode(COL.FILE_NAME, Cnt)
            !RelationalCode = PublishCode(COL.RELATIONAL_CODE, Cnt)
            !ProductCode = PublishCode(COL.PRODUCT_CODE, Cnt)
            .Update
        Next Cnt
        'オブジェクトを閉じる
        .Close
    End With
    
    'オブジェクト解放
    Set TargetTable = Nothing
End Sub

' ----------------------------------------------------------------------
'   関数名：GetProductData（商品情報取得処理）
'   引数1： CurPath（カレントフォルダのパス）
'   引数2： CurDate（現在の日付）
'   引数3： ConvFilePath（CSVファイルがエクセルファイルに変換された後のパスを受け取る変数）
'   戻り値：なし
' ----------------------------------------------------------------------
Sub GetProductData(CurPath As String, CurDate As String, ConvFilePath As String)
    '変数宣言
    Dim ExApp As Excel.Application
    Dim TargetBook As Workbook
    Dim CopyFilePath As String
    
    '商品情報ファイルをコピー
    CopyFilePath = CurPath & COPY_FILE_NAME
    If (Dir(CopyFilePath) <> "") Then Kill CopyFilePath
    FileCopy Forms!F_StartUp!TextBox10.Value, CopyFilePath
    
    'エクセルオブジェクトを取得
    Set ExApp = CreateObject("Excel.application")
    
    'コピーしたCSVファイルを開く
    Set TargetBook = ExApp.Workbooks.Open(CopyFilePath)
    
    'コピーしたCSVファイルをエクセルファイルで保存
    ConvFilePath = CurPath & "\" & LINK_FILE_NAME & "_" & CurDate & EXT_EXCEL
    If (Dir(ConvFilePath) <> "") Then Kill ConvFilePath
    With TargetBook
        .SaveAs FileName:=ConvFilePath, FileFormat:=xlOpenXMLWorkbook
        .Close
    End With
    
    'コピーしたCSVファイルを削除
    Kill CopyFilePath
    
    'オブジェクト解放
    Set TargetBook = Nothing
    ExApp.Quit
    Set ExApp = Nothing
End Sub

' ----------------------------------------------------------------------
'   関数名：ExportBook（商品情報をエクセルファイル出力する処理）
'   引数1： CurPath（カレントフォルダのパス）
'   引数2： CurDate（現在の日付）
'   引数3： ConvFilePath（CSVファイルがエクセルファイルに変換された後のパス）
'   引数4： ExpFilePath（出力するエクセルファイルのパスを受け取る変数）
'   戻り値：Bollean（正常=True/異常=False）
' ----------------------------------------------------------------------
Function ExportBook(CurPath As String, CurDate As String, ConvFilePath As String, ExpFilePath As String) As Boolean
    '変数宣言
    Dim ExApp As Excel.Application
    Dim TargetBook As Workbook
    
    '関数の戻り値を設定
    ExportBook = True
    
    'オブジェクトを取得
    Set ExApp = CreateObject("Excel.application")
    Set TargetBook = ExApp.Workbooks.Open(ConvFilePath)
    
    'エクセルファイルの書式を編集（文字列）
    If (SettingFormat(TargetBook, WORD_FORM_STR, PROC_FLG.F_STR) = False) Then
        MsgBox NOT_FOUND_STR_COL, vbOKOnly, MSG_TITLE
        ExportBook = False
        Exit Function
    End If
    
    'エクセルファイルの書式を編集（数値）
    If (SettingFormat(TargetBook, WORD_FORM_INT, PROC_FLG.F_INT) = False) Then
        MsgBox NOT_FOUND_INT_COL, vbOKOnly, MSG_TITLE
        ExportBook = False
        Exit Function
    End If
    
    'エクセルファイルの書式を編集（日付）
    If (SettingFormat(TargetBook, WORD_FORM_DATE, PROC_FLG.F_DATE) = False) Then
        MsgBox NOT_FOUND_DATE_COL, vbOKOnly, MSG_TITLE
        ExportBook = False
        Exit Function
    End If
    
    'ファイルを保存して閉じる
    With TargetBook
        .Sheets(TOP_SHEET).Cells(1, 1).Select
        .Close SaveChanges:=True
    End With
    
    'エクセルファイルに対するリンクテーブルを作成
    DoCmd.TransferSpreadsheet acLink, acSpreadsheetTypeExcel12Xml, LINK_FILE_NAME, ConvFilePath, True

    'クエリ実行（ローカルテーブル作成）
    DoCmd.OpenQuery QUERY_C10
    
    'リンクテーブル解除し、エクセルファイルを削除
    DoCmd.DeleteObject acTable, LINK_FILE_NAME
    Kill ConvFilePath
    
    'クエリをエクセルファイルで出力
    ExpFilePath = CurPath & "\" & EXP_FILE_NAME & "_" & CurDate & EXT_EXCEL
    If (Dir(ExpFilePath) <> "") Then Kill ExpFilePath
    DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel12Xml, EXP_FILE_NAME, ExpFilePath, True
    DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel12Xml, EXP_DATA_NAME, ExpFilePath, True
    
    'オブジェクト解放
    Set TargetBook = Nothing
    ExApp.Quit
    Set ExApp = Nothing
End Function

' ----------------------------------------------------------------------
'   関数名：SettingFormat（リンクテーブルとなるエクセルファイルの書式編集処理）
'   引数1： TargetBook（編集対象のエクセルファイル）
'   引数2： WordList（カンマ区切りのカラム名）
'   引数3： FormatTypeFlg（列挙型PROC_FLGの書式種別フラグ）
'   戻り値：Bollean（正常=True/異常=False）
' ----------------------------------------------------------------------
Function SettingFormat(TargetBook As Workbook, WordList As String, FormatTypeFlg As Integer) As Boolean
    '変数宣言
    Dim TargetSheet As Worksheet
    Dim AllRowRange As Range
    Dim HeaderRange As Range
    Dim TargetCell As Range
    Dim StartCell As Range
    Dim EndCell As Range
    Dim TargetWord As Variant
    
    '関数の戻り値を設定
    SettingFormat = True
    
    'オブジェクト取得
    Set TargetSheet = TargetBook.Sheets(TOP_SHEET)
    Set AllRowRange = TargetSheet.Rows
    
    With TargetSheet
        'ヘッダーの範囲を取得
        Set HeaderRange = .Range(.Cells(1, 1), .Cells(1, 1).End(xlToRight))

        For Each TargetWord In Split(WordList, ",")
            'ヘッダー範囲を検索
            Set TargetCell = HeaderRange.Find(What:=TargetWord, LookAt:=xlWhole)
            If TargetCell Is Nothing Then
                SettingFormat = False
                Exit Function
            End If
            
            '書式を設定するセルを取得
            Set StartCell = .Cells(TargetCell.Offset(1, 0).Row, TargetCell.Column)
            Set EndCell = .Cells(.Cells(AllRowRange.Count, 1).End(xlUp).Row, TargetCell.Column)

            '書式設定
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
    
    'オブジェクト解放
    Set TargetSheet = Nothing
    Set AllRowRange = Nothing
    Set HeaderRange = Nothing
    Set TargetCell = Nothing
    Set StartCell = Nothing
    Set EndCell = Nothing
End Function

' ----------------------------------------------------------------------
'   関数名：EditBook（出力エクセルファイル編集処理）
'   引数： ExpFilePath（出力したエクセルファイルのパス）
'   戻り値：なし
' ----------------------------------------------------------------------
Sub EditBook(ExpFilePath As String)
    '変数宣言
    Dim ExApp As Excel.Application
    Dim TargetBook As Workbook
    Dim TargetSheet As Worksheet
    Dim HeaderRange As Range
    
    'エクセルオブジェクトを取得
    Set ExApp = CreateObject("Excel.application")

    Set TargetBook = ExApp.Workbooks.Open(ExpFilePath)
    For Each TargetSheet In TargetBook.Sheets
        With TargetSheet
            .Select
            .Cells(1, 1).Select
            'ヘッダーの範囲を取得
            Set HeaderRange = .Range(.Cells(1, 1), .Cells(1, 1).End(xlToRight))
        End With
        'ヘッダーの設定（中央揃え/列幅自動設定）
        With HeaderRange
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .AutoFilter
            .Columns.EntireColumn.AutoFit
        End With
        '表示画面の設定（80%縮小/ヘッダー固定）
        With ExApp.ActiveWindow
            .Zoom = ZOOM_POINT
            .SplitRow = HEADER_ROW
            .FreezePanes = True
        End With
    Next
    
    'ファイルを上書き保存して閉じる
    With TargetBook
        .Sheets(TOP_SHEET).Select
        .Save
        .Close
    End With
    
    'オブジェクト解放
    Set TargetBook = Nothing
    ExApp.Quit
    Set ExApp = Nothing
End Sub
